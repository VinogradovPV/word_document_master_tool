def test_gui_app_imports() -> None:
    from word_document_master_tool.gui.app import main

    assert callable(main)


def test_main_window_imports() -> None:
    from word_document_master_tool.gui.main_window import MainWindow

    assert MainWindow is not None
