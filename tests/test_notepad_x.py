import base64
import importlib.util
import json
import os
from pathlib import Path
import tempfile
import threading
import time
import unittest
from unittest import mock


MODULE_PATH = Path(__file__).resolve().parents[1] / 'Notepad-X.py'
SPEC = importlib.util.spec_from_file_location('notepad_x', MODULE_PATH)
notepad_x = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(notepad_x)


class FakeText:
    def __init__(self, content=''):
        self.content = content

    def get(self, start, end):
        if start == '1.0' and end == 'end-1c':
            return self.content
        return self.content


class FakeEditableText(FakeText):
    def __init__(self, content='', modified=False):
        super().__init__(content)
        self.modified = modified

    def edit_modified(self, value=None):
        if value is not None:
            self.modified = bool(value)
        return self.modified


def make_bare_app():
    app = notepad_x.NotepadX.__new__(notepad_x.NotepadX)
    app.tr = lambda _key, default, **values: default.format(**values)
    return app


class LargeFileIndexTests(unittest.TestCase):
    def test_line_lengths_are_not_double_counted(self):
        payload = b'a\n' + (b'b' * 70_000) + b'\nccc'
        with tempfile.NamedTemporaryFile(delete=False) as source:
            source.write(payload)
            source_path = source.name
        try:
            ranges = ((0, 20_000), (20_000, 50_000), (50_000, len(payload)))
            lengths = []
            for index, (start, end) in enumerate(ranges):
                result = notepad_x.scan_large_text_file_index_range_worker(
                    source_path, start, end, chunk_size=64 * 1024, range_index=index
                )
                lengths.extend(result['line_lengths'])
            self.assertEqual([1, 70_000, 3], lengths)
        finally:
            os.remove(source_path)

    def test_trailing_newline_produces_an_empty_final_line(self):
        with tempfile.NamedTemporaryFile(delete=False) as source:
            source.write(b'one\ntwo\n')
            source_path = source.name
        try:
            result = notepad_x.scan_large_text_file_index_range_worker(
                source_path, 0, os.path.getsize(source_path)
            )
            self.assertEqual([3, 3, 0], result['line_lengths'])
        finally:
            os.remove(source_path)

    def test_grouped_index_output_is_bounded_and_preserves_line_count(self):
        payload = (b'x\n' * 10_000) + b'last'
        with tempfile.NamedTemporaryFile(delete=False) as source:
            source.write(payload)
            source_path = source.name
        try:
            result = notepad_x.scan_large_text_file_index_range_worker(
                source_path,
                0,
                len(payload),
                line_group_size=256
            )
            self.assertEqual(10_001, result['total_lines'])
            self.assertEqual(result['total_lines'], sum(result['line_group_counts']))
            self.assertLessEqual(len(result['line_group_max_lengths']), 40)
            self.assertEqual(len(result['line_group_counts']), len(result['line_group_max_lengths']))
        finally:
            os.remove(source_path)

    def test_progressive_minimap_grows_beyond_initial_one_line_hint(self):
        app = make_bare_app()
        app.minimap_max_segments = 240
        doc = {'frame': 'tab-1'}
        app.start_progressive_minimap_build(doc, total_lines_hint=None)
        first_chunk = ('short\n' * 400) + 'partial'
        doc['background_lines_loaded'] = 401
        app.append_progressive_minimap_chunk(doc, first_chunk, finalize=False)
        doc['background_lines_loaded'] = 402
        app.append_progressive_minimap_chunk(doc, '-line\nlast', finalize=True)

        model = doc['minimap_model']
        self.assertEqual(402, model['total_lines'])
        self.assertTrue(model['complete'])
        self.assertLessEqual(len(model['segments']), app.minimap_max_segments)

    def test_dense_virtual_index_stops_at_configured_limit(self):
        with tempfile.NamedTemporaryFile(delete=False) as source:
            source.write(b'x\n' * 100)
            source_path = source.name
        try:
            with self.assertRaises(RuntimeError):
                notepad_x.scan_newline_start_offsets_range_worker(
                    source_path,
                    0,
                    os.path.getsize(source_path),
                    max_line_starts=10,
                )
        finally:
            os.remove(source_path)

    def test_virtual_line_index_replacement_does_not_duplicate_suffix_boundary(self):
        app = make_bare_app()
        doc = {'line_starts': [0, 2, 4, 6, 8], 'total_file_lines': 5, 'file_size_bytes': 9}

        app.update_virtual_line_index_after_replace(doc, 2, 3, 2, 6, b'b\nc\n')

        self.assertEqual([0, 2, 4, 6, 8], doc['line_starts'])

    def test_oversized_virtual_line_is_bounded_and_read_only(self):
        app = make_bare_app()
        app.virtual_file_window_max_bytes = 32
        app.huge_file_preview_threshold_bytes = 1024
        app.huge_virtual_file_window_max_bytes = 16
        with tempfile.NamedTemporaryFile(delete=False) as source:
            source.write(b'a' * 100)
            source_path = source.name
        try:
            doc = {
                'file_path': source_path,
                'line_starts': [0],
                'total_file_lines': 1,
                'file_size_bytes': 100,
                'virtual_mode': True,
                'virtual_editable': False,
            }
            content = app.read_virtual_line_window(doc, 1, 1)
            self.assertEqual(32, len(content))
            self.assertTrue(doc['virtual_window_truncated'])
            self.assertTrue(app.is_doc_text_readonly(doc))
        finally:
            os.remove(source_path)


class TextIntegrityTests(unittest.TestCase):
    def test_widget_content_preserves_all_trailing_newlines(self):
        app = make_bare_app()
        self.assertEqual('body\n\n', app.get_text_widget_content(FakeText('body\n\n')))

    def test_replace_treats_backslashes_as_literal_text(self):
        app = make_bare_app()
        replacement = r'\g<missing>\folder'
        result, count = app.replace_query_in_text('token\n\n', 'token', replacement)
        self.assertEqual(1, count)
        self.assertEqual(replacement + '\n\n', result)

    def test_replace_with_no_match_does_not_change_content(self):
        app = make_bare_app()
        result, count = app.replace_query_in_text('original\n', 'missing', 'value')
        self.assertEqual(0, count)
        self.assertEqual('original\n', result)


class PersistenceSafetyTests(unittest.TestCase):
    def test_clean_frozen_build_uses_user_support_directory(self):
        app = make_bare_app()
        with tempfile.TemporaryDirectory() as directory:
            executable_dir = Path(directory) / 'release'
            support_dir = Path(directory) / 'user-state'
            executable_dir.mkdir()
            support_dir.mkdir()
            app.get_user_support_dir = lambda: str(support_dir)
            app.get_emergency_support_dir = lambda: str(Path(directory) / 'emergency')

            with (
                mock.patch.object(notepad_x.sys, 'frozen', True, create=True),
                mock.patch.object(notepad_x.sys, 'executable', str(executable_dir / 'Notepad-X.exe')),
            ):
                self.assertEqual(str(support_dir), app.get_app_dir())

    def test_existing_frozen_portable_config_keeps_executable_directory(self):
        app = make_bare_app()
        with tempfile.TemporaryDirectory() as directory:
            executable_dir = Path(directory) / 'portable'
            (executable_dir / 'cfg').mkdir(parents=True)
            app.get_user_support_dir = lambda: str(Path(directory) / 'user-state')
            app.get_emergency_support_dir = lambda: str(Path(directory) / 'emergency')

            with (
                mock.patch.object(notepad_x.sys, 'frozen', True, create=True),
                mock.patch.object(notepad_x.sys, 'executable', str(executable_dir / 'Notepad-X.exe')),
            ):
                self.assertEqual(str(executable_dir), app.get_app_dir())

    def test_external_change_and_deletion_are_detected(self):
        app = make_bare_app()
        with tempfile.NamedTemporaryFile(delete=False) as source:
            source.write(b'first')
            source_path = source.name
        try:
            doc = {
                'file_path': source_path,
                'file_signature': app.get_file_signature(source_path),
                'is_remote': False,
            }
            with open(source_path, 'wb') as changed:
                changed.write(b'second')
            self.assertTrue(app.doc_file_changed_on_disk(doc))
            doc['file_signature'] = app.get_file_signature(source_path)
            os.remove(source_path)
            self.assertTrue(app.doc_file_changed_on_disk(doc))
        finally:
            if os.path.exists(source_path):
                os.remove(source_path)

    def test_atomic_text_write_does_not_fall_back_to_truncating_target(self):
        app = make_bare_app()
        app.is_windows = True
        app.log_exception = lambda *_args, **_kwargs: None
        with tempfile.TemporaryDirectory() as directory:
            target = os.path.join(directory, 'document.txt')
            Path(target).write_text('original', encoding='utf-8')
            with mock.patch.object(notepad_x.os, 'replace', side_effect=PermissionError('locked')):
                with self.assertRaises(PermissionError):
                    app.write_file_atomically(target, 'replacement')
            self.assertEqual('original', Path(target).read_text(encoding='utf-8'))

    def test_atomic_json_write_preserves_target_when_replace_fails(self):
        app = make_bare_app()
        app.is_windows = True
        app.log_exception = lambda *_args, **_kwargs: None
        with tempfile.TemporaryDirectory() as directory:
            target = os.path.join(directory, 'session.json')
            Path(target).write_text('{"state":"original"}', encoding='utf-8')
            with mock.patch.object(notepad_x.os, 'replace', side_effect=PermissionError('locked')):
                written = app.write_json_atomically(target, {'state': 'new'}, 'test-', 'test write')
            self.assertFalse(written)
            self.assertEqual({'state': 'original'}, json.loads(Path(target).read_text(encoding='utf-8')))

    def test_json_reader_rejects_oversized_support_file(self):
        app = make_bare_app()
        app.log_exception = lambda *_args, **_kwargs: None
        with tempfile.NamedTemporaryFile(delete=False) as source:
            source.write(b'{"value":"' + (b'x' * 100) + b'"}')
            source_path = source.name
        try:
            result = app.read_json_file(source_path, 'test read', {'fallback': True}, max_bytes=32)
            self.assertEqual({'fallback': True}, result)
        finally:
            os.remove(source_path)

    def test_same_file_detection_handles_aliases(self):
        app = make_bare_app()
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'document.txt'
            target.write_text('content', encoding='utf-8')
            alias = Path(directory) / '.' / 'document.txt'
            self.assertTrue(app.paths_refer_to_same_file(str(target), str(alias)))

    def test_failed_save_as_restores_document_metadata(self):
        app = make_bare_app()
        old_path = os.path.abspath('old.npxe')
        new_path = os.path.abspath('new.txt')
        doc = {
            'frame': 'tab-1',
            'file_path': old_path,
            'file_signature': (1, 2),
            'is_remote': False,
            'remote_spec': None,
            'remote_host': None,
            'remote_path': None,
            'remote_shadow_path': None,
            'display_name': None,
            'untitled_name': None,
            'encrypted_file': True,
            'encryption_header': {'header': True},
            'encryption_key': b'secret',
            'notes_registered': True,
            'preview_mode': False,
            'virtual_mode': False,
        }
        app.get_current_doc = lambda: doc
        app.prompt_plain_text_save_path = lambda _title: new_path
        app.paths_refer_to_same_file = lambda _first, _second: False
        app.unregister_doc_from_shared_notes = lambda current: current.update(notes_registered=False)
        app.register_doc_for_shared_notes = lambda current: current.update(notes_registered=True)
        app.clear_remote_metadata = lambda current: current.update(
            is_remote=False, remote_spec=None, remote_host=None, remote_path=None, remote_shadow_path=None
        )
        app.get_file_signature = lambda _path: None
        app.configure_syntax_highlighting = lambda _frame: None
        app.set_active_document = lambda _frame: None
        app.refresh_tab_title = lambda _frame: None
        app.save_document_content = lambda *_args, **_kwargs: False

        self.assertFalse(app.save_as())
        self.assertEqual(old_path, doc['file_path'])
        self.assertTrue(doc['encrypted_file'])
        self.assertEqual(b'secret', doc['encryption_key'])
        self.assertTrue(doc['notes_registered'])

    def test_load_signature_does_not_adopt_a_newer_external_version(self):
        app = make_bare_app()
        with tempfile.NamedTemporaryFile(delete=False) as source:
            source.write(b'first')
            source_path = source.name
        try:
            original_signature = app.get_file_signature(source_path)
            doc = {'file_path': source_path, 'load_source_signature': original_signature}
            Path(source_path).write_bytes(b'a newer external version')

            self.assertFalse(app.finalize_doc_file_signature_after_load(doc))
            self.assertEqual(original_signature, doc['file_signature'])
            self.assertTrue(doc['autosave_conflict'])
            self.assertTrue(app.doc_file_changed_on_disk({**doc, 'is_remote': False}))
        finally:
            os.remove(source_path)

    @unittest.skipIf(os.name == 'nt', 'Creating file symlinks may require Windows developer mode')
    def test_atomic_save_preserves_a_symlink(self):
        app = make_bare_app()
        app.is_windows = False
        app.log_exception = lambda *_args, **_kwargs: None
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'target.txt'
            link = Path(directory) / 'link.txt'
            target.write_text('old', encoding='utf-8')
            link.symlink_to(target)

            app.write_file_atomically(str(link), 'new')

            self.assertTrue(link.is_symlink())
            self.assertEqual('new', target.read_text(encoding='utf-8'))

    def test_atomic_copy_preserves_existing_destination_on_replace_failure(self):
        app = make_bare_app()
        app.is_windows = True
        app.file_load_chunk_size = 4
        app.log_exception = lambda *_args, **_kwargs: None
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'source.bin'
            target = Path(directory) / 'target.bin'
            source.write_bytes(b'replacement')
            target.write_bytes(b'original')
            with mock.patch.object(notepad_x.os, 'replace', side_effect=PermissionError('locked')):
                with self.assertRaises(PermissionError):
                    app.copy_file_atomically(str(source), str(target))
            self.assertEqual(b'original', target.read_bytes())


class AsyncAndStateValidationTests(unittest.TestCase):
    def test_loading_document_is_read_only_and_cannot_be_saved(self):
        app = make_bare_app()
        app.root = object()
        doc = {'loading_file': True, 'background_loading': True}
        self.assertTrue(app.is_doc_text_readonly(doc))
        with mock.patch.object(notepad_x.messagebox, 'showinfo'):
            self.assertFalse(app.save_document_content(doc, show_errors=True))

    def test_stale_background_chunk_is_ignored_after_token_is_cleared(self):
        app = make_bare_app()
        source_path = os.path.abspath('large.txt')
        doc = {'file_path': source_path, 'background_load_token': None, 'background_index_token': None}
        app.documents = {'tab-1': doc}
        app.append_background_text_chunk = mock.Mock()

        app.handle_background_file_result({
            'kind': 'text_chunk',
            'tab_id': 'tab-1',
            'token': 'old-token',
            'file_path': source_path,
            'chunk_text': 'stale',
        })

        app.append_background_text_chunk.assert_not_called()

    def test_matching_background_chunk_is_dispatched(self):
        app = make_bare_app()
        source_path = os.path.abspath('large.txt')
        doc = {'file_path': source_path, 'background_load_token': 'current', 'background_index_token': None}
        app.documents = {'tab-1': doc}
        app.append_background_text_chunk = mock.Mock()

        app.handle_background_file_result({
            'kind': 'text_chunk',
            'tab_id': 'tab-1',
            'token': 'current',
            'file_path': source_path,
            'chunk_text': 'current data',
        })

        app.append_background_text_chunk.assert_called_once()

    def test_scalar_session_and_recovery_collections_are_sanitized(self):
        app = make_bare_app()
        app.max_session_files = 100
        app.max_recent_files = 10
        app.max_search_history = 10
        app.max_search_history_entry_length = 512
        app.max_command_history = 25
        app.max_command_history_entry_length = 1000
        app.base_font_size = 13
        app.min_font_size = 6
        app.max_font_size = 32
        app.locale_code = 'en_us'
        app.command_panel_default_height = 8
        app.command_panel_min_height = 5
        app.command_panel_max_height = 28
        app.default_window_width = 900
        app.default_window_height = 700
        app.get_available_syntax_theme_names = lambda: ['Default']
        app.normalize_locale_code = lambda value: 'en_us'
        app.get_locale_file_path = lambda *_args, **_kwargs: __file__
        app.sanitize_hotkey_overrides = lambda _payload: {}
        session = app.sanitize_session_payload({
            'open_files': 1,
            'recent_files': None,
            'closed_session_files': {'bad': 'shape'},
        })
        self.assertEqual([], session['open_files'])
        self.assertEqual([], session['recent_files'])
        self.assertEqual([], session['closed_session_files'])

        app.max_recovery_tabs = 20
        app.next_untitled_name = lambda: 'Untitled 1'
        recovery = app.sanitize_recovery_payload({'recovery_tabs': 1})
        self.assertEqual([], recovery['recovery_tabs'])

    def test_shutdown_uses_process_pool_termination_when_available(self):
        app = make_bare_app()
        executor = mock.Mock()
        app.index_process_executor = executor
        app.shutdown_index_process_executor()
        executor.terminate_workers.assert_called_once_with()
        executor.shutdown.assert_not_called()


class SecurityBoundaryTests(unittest.TestCase):
    def test_network_path_is_rejected_without_filesystem_probe(self):
        with mock.patch.object(notepad_x.os.path, 'exists') as exists:
            result = notepad_x.normalize_startup_path_argument(
                r'\\server\share\secret.txt',
                allow_network_paths=False,
            )
        self.assertIsNone(result)
        exists.assert_not_called()

    def test_authenticated_single_instance_round_trip(self):
        with tempfile.TemporaryDirectory() as app_dir, tempfile.TemporaryDirectory() as state_dir:
            source_path = os.path.join(app_dir, 'open-me.txt')
            Path(source_path).write_text('content', encoding='utf-8')
            with mock.patch.dict(os.environ, {'LOCALAPPDATA': state_dir}):
                app = make_bare_app()
                app.isolated_session = False
                app.is_windows = True
                app.single_instance_secret = notepad_x.load_or_create_notepadx_ipc_secret(app_dir)
                app.single_instance_host = '127.0.0.1'
                app.single_instance_port = notepad_x.get_notepadx_single_instance_port(app_dir)
                app.single_instance_server = None
                app.single_instance_running = False
                app.single_instance_listener_thread = None
                app.remote_open_files = []
                app.remote_open_lock = threading.Lock()
                app.start_single_instance_server()
                try:
                    self.assertTrue(notepad_x.send_files_to_running_notepadx(app_dir, [source_path]))
                    deadline = time.monotonic() + 1.0
                    while not app.remote_open_files and time.monotonic() < deadline:
                        time.sleep(0.01)
                    self.assertEqual([os.path.abspath(source_path)], app.remote_open_files)
                finally:
                    app.stop_single_instance_server()

    def test_rogue_single_instance_listener_receives_no_file_paths(self):
        with tempfile.TemporaryDirectory() as app_dir, tempfile.TemporaryDirectory() as state_dir:
            source_path = os.path.join(app_dir, 'private-name.txt')
            Path(source_path).write_text('content', encoding='utf-8')
            captured = []
            ready = threading.Event()
            port = notepad_x.get_notepadx_single_instance_port(app_dir)

            def rogue_server():
                with notepad_x.socket.socket(notepad_x.socket.AF_INET, notepad_x.socket.SOCK_STREAM) as server:
                    if hasattr(notepad_x.socket, 'SO_EXCLUSIVEADDRUSE'):
                        server.setsockopt(notepad_x.socket.SOL_SOCKET, notepad_x.socket.SO_EXCLUSIVEADDRUSE, 1)
                    server.bind(('127.0.0.1', port))
                    server.listen(1)
                    ready.set()
                    connection, _address = server.accept()
                    with connection:
                        connection.settimeout(1.0)
                        stream = connection.makefile('rwb')
                        captured.append(stream.readline(notepad_x.SINGLE_INSTANCE_MAX_PAYLOAD_BYTES + 2))
                        notepad_x.write_notepadx_ipc_message(stream, {
                            'protocol': notepad_x.SINGLE_INSTANCE_PROTOCOL,
                            'type': 'challenge',
                            'server_nonce': notepad_x.encode_notepadx_ipc_nonce(
                                b'0' * notepad_x.SINGLE_INSTANCE_NONCE_BYTES
                            ),
                            'proof': '00' * 32,
                        })
                        captured.append(stream.readline(notepad_x.SINGLE_INSTANCE_MAX_PAYLOAD_BYTES + 2))
                        stream.close()

            listener = threading.Thread(target=rogue_server, daemon=True)
            listener.start()
            self.assertTrue(ready.wait(1.0))
            with mock.patch.dict(os.environ, {'LOCALAPPDATA': state_dir}):
                self.assertFalse(notepad_x.send_files_to_running_notepadx(app_dir, [source_path]))
            listener.join(1.0)
            transmitted = b''.join(captured)
            self.assertNotIn(os.path.basename(source_path).encode('utf-8'), transmitted)
            self.assertEqual(b'', captured[-1])

    def test_corrupt_ipc_secret_fails_closed(self):
        with tempfile.TemporaryDirectory() as app_dir, tempfile.TemporaryDirectory() as state_dir:
            with mock.patch.dict(os.environ, {'LOCALAPPDATA': state_dir}):
                secret_path = notepad_x.get_notepadx_ipc_secret_path(app_dir)
                os.makedirs(os.path.dirname(secret_path), exist_ok=True)
                Path(secret_path).write_bytes(b'short')
                self.assertIsNone(notepad_x.load_or_create_notepadx_ipc_secret(app_dir))
                self.assertFalse(notepad_x.send_files_to_running_notepadx(app_dir, [__file__]))

    def test_owned_remote_cleanup_cannot_delete_an_unowned_file(self):
        app = make_bare_app()
        app.log_exception = lambda *_args, **_kwargs: None
        with tempfile.TemporaryDirectory() as directory:
            app.remote_cache_dir = os.path.join(directory, 'cache')
            os.makedirs(app.remote_cache_dir)
            app.owned_remote_shadow_paths = set()
            outside_path = os.path.join(directory, 'keep.txt')
            Path(outside_path).write_text('keep', encoding='utf-8')
            self.assertFalse(app.cleanup_owned_remote_shadow(outside_path))
            self.assertTrue(os.path.exists(outside_path))

    def test_frozen_build_never_relaunches_itself_as_pip(self):
        app = make_bare_app()
        app.is_windows = True
        with mock.patch.object(notepad_x.sys, 'frozen', True, create=True):
            self.assertIsNone(app.get_pip_python_executable())


class EncryptionValidationTests(unittest.TestCase):
    def setUp(self):
        self.app = make_bare_app()
        self.app.encryption_magic = b'NPXENC1'
        self.app.encryption_version = 1
        self.app.encryption_nonce_length = 12
        self.app.encryption_scrypt_n = 1 << 15
        self.app.encryption_scrypt_r = 8
        self.app.encryption_scrypt_p = 1
        self.app.encryption_scrypt_maxmem = 128 * 1024 * 1024
        self.app.encryption_scrypt_max_work_factor = (1 << 15) * 8 * 4
        self.app.encryption_header_max_bytes = 64 * 1024

    def valid_header(self):
        return {
            'format': 'Notepad-X Encrypted',
            'version': 1,
            'cipher': 'AES-256-GCM',
            'kdf': 'scrypt',
            'n': 1 << 15,
            'r': 8,
            'p': 1,
            'salt': base64.b64encode(b'0' * 16).decode('ascii'),
            'encoding': 'utf-8',
        }

    def test_rejects_excessive_scrypt_work(self):
        header = self.valid_header()
        header['n'] = 1 << 22
        with self.assertRaises(ValueError):
            self.app.validate_encryption_header(header)

    def test_rejects_invalid_base64_salt(self):
        header = self.valid_header()
        header['salt'] = '***not-base64***'
        with self.assertRaises(ValueError):
            self.app.validate_encryption_header(header)

    def test_rejects_oversized_header_before_parsing_json(self):
        payload = (
            self.app.encryption_magic
            + (self.app.encryption_header_max_bytes + 1).to_bytes(4, 'big')
            + b'{}'
        )
        with self.assertRaises(ValueError):
            self.app.parse_encrypted_payload(payload)

    def test_recovery_state_never_serializes_encrypted_plaintext(self):
        encrypted_doc = {
            'encrypted_file': True,
            'large_file_mode': False,
            'file_path': 'secret.npxe',
            'text': FakeText('plaintext secret'),
        }
        self.app.documents = {'tab-1': encrypted_doc}
        self.app.get_current_doc = lambda: encrypted_doc
        self.app.max_recovery_tabs = 20
        self.app.utc_timestamp = lambda: '2026-01-01T00:00:00+00:00'

        recovery = self.app.get_recovery_state()

        self.assertEqual([], recovery['recovery_tabs'])


class ProjectAndDispatchTests(unittest.TestCase):
    def test_project_scan_honors_file_limit(self):
        app = make_bare_app()
        with tempfile.TemporaryDirectory() as directory:
            selected = Path(directory) / 'selected.py'
            selected.write_text('pass\n', encoding='utf-8')
            for index in range(20):
                (Path(directory) / f'module_{index:02}.py').write_text('pass\n', encoding='utf-8')
            files = app.get_project_source_files(str(selected), max_files=5)
            self.assertEqual(5, len(files))
            self.assertEqual(os.path.abspath(selected), files[0])

    def test_grab_git_scan_honors_limit_and_regular_files_only(self):
        app = make_bare_app()
        app.grab_git_max_project_files = 3
        with tempfile.TemporaryDirectory() as directory:
            for index in range(8):
                (Path(directory) / f'module_{index:02}.py').write_text('pass\n', encoding='utf-8')
            files = app.get_grab_git_project_files(directory)
            self.assertEqual(3, len(files))
            self.assertTrue(all(Path(path).is_file() for path in files))

    def test_command_completion_is_dispatched_without_document_lookup(self):
        app = make_bare_app()
        received = []
        app.finish_shell_command = lambda command, result: received.append((command, result))
        app.documents = {}
        payload = {'kind': 'command_complete', 'command_text': 'echo ok', 'result': {'returncode': 0}}
        app.handle_background_file_result(payload)
        self.assertEqual([('echo ok', {'returncode': 0})], received)

    def test_find_in_files_stops_at_configured_limits(self):
        app = make_bare_app()
        app.find_in_max_file_bytes = 1024
        app.find_in_max_files = 2
        app.find_in_max_results = 10
        app.find_in_max_total_matches = 100
        app.get_find_in_supported_patterns = lambda: (set(), {'.txt'})
        with tempfile.TemporaryDirectory() as directory:
            for index in range(3):
                (Path(directory) / f'document_{index}.txt').write_text('needle\n', encoding='utf-8')

            results, scanned_files, total_matches, limited = app.search_query_in_directory(
                directory,
                'needle',
            )

        self.assertEqual(2, scanned_files)
        self.assertEqual(2, total_matches)
        self.assertEqual(2, len(results))
        self.assertTrue(limited)

    def test_stop_command_runner_terminates_owned_process(self):
        app = make_bare_app()
        app.log_exception = lambda *_args, **_kwargs: None
        process = mock.Mock()
        process.poll.return_value = None
        process.wait.return_value = 0
        app.command_runner_process = process

        app.stop_command_runner()

        process.terminate.assert_called_once_with()
        process.wait.assert_called_once_with(timeout=0.75)
        self.assertIsNone(app.command_runner_process)


class SharedNoteValidationTests(unittest.TestCase):
    def test_non_numeric_note_id_does_not_crash_tag_creation(self):
        app = make_bare_app()
        app.note_colors = {'yellow': '#fff000'}
        app.max_note_text_length = 4000
        app.max_note_name_length = 120
        app.apply_note_tag = lambda *_args, **_kwargs: None
        doc = {'note_counter': 1, 'notes': {}, 'text': FakeText('anchor')}

        note_tag = app.create_note_tag(doc, '1.0', '1.6', 'Review this', note_id='external-note-id')

        self.assertTrue(note_tag.startswith('note_'))
        self.assertEqual('external-note-id', doc['notes'][note_tag]['id'])

    def test_control_characters_in_note_id_are_replaced(self):
        app = make_bare_app()
        app.note_colors = {'yellow': '#fff000'}
        app.max_note_text_length = 4000
        app.max_note_name_length = 120
        app.apply_note_tag = lambda *_args, **_kwargs: None
        doc = {'note_counter': 7, 'notes': {}, 'text': FakeText('anchor')}

        note_tag = app.create_note_tag(doc, '1.0', '1.6', 'Review this', note_id='bad\x01id')

        self.assertEqual('note_7', note_tag)
        self.assertEqual('7', doc['notes'][note_tag]['id'])

    def test_remote_pid_is_not_checked_against_local_process_table(self):
        app = make_bare_app()
        app.max_shared_editors = 32
        app.max_shared_editor_id_length = 128
        app.max_shared_editor_label_length = 80
        app.max_shared_editor_host_length = 128
        app.max_shared_editor_ip_length = 64
        app.shared_editor_stale_seconds = 30
        app.get_local_machine_name = lambda: 'local-host'
        editors = [{
            'id': 'remote-editor',
            'label': 'Notepad-X-2',
            'pid': os.getpid(),
            'last_seen': '2000-01-01T00:00:00+00:00',
            'host': 'remote-host',
            'ip': '192.0.2.10',
        }]

        self.assertEqual([], app.prune_inactive_shared_editors(editors))

    def test_case_distinct_sidecars_are_not_merged_on_case_sensitive_platforms(self):
        app = make_bare_app()
        app.is_windows = False
        with tempfile.TemporaryDirectory() as directory:
            upper = Path(directory) / 'File.py.notepadx.notes.json'
            lower = Path(directory) / 'file.py.notepadx.notes.json'
            upper.write_text('{"notes": []}', encoding='utf-8')
            lower.write_text('{"notes": []}', encoding='utf-8')

            self.assertEqual([str(upper)], app.get_sidecar_variants(str(upper)))
            app.cleanup_duplicate_sidecar_variants(str(upper))
            self.assertTrue(upper.exists())
            self.assertTrue(lower.exists())


if __name__ == '__main__':
    unittest.main()
