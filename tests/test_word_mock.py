import os
import sys
import unittest
from unittest.mock import MagicMock, patch

# Добавляем путь к src, чтобы импорты работали
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from word_document_master_tool.word.word_app import WordApp


class TestWordAppMock(unittest.TestCase):
    @patch('word_document_master_tool.word.word_app.win32com')
    @patch('word_document_master_tool.word.word_app.pythoncom')
    @patch('platform.system')
    def test_word_app_lifecycle(self, mock_system, mock_pythoncom, mock_win32com):
        # Настраиваем моки
        mock_system.return_value = "Windows"
        mock_app = MagicMock()
        mock_win32com.client.DispatchEx.return_value = mock_app
        mock_app.Documents.Count = 0
        
        # Тестируем контекстный менеджер
        with WordApp() as word:
            self.assertEqual(word.app, mock_app)
            self.assertTrue(word._owns_app)
            
        # Проверяем вызовы
        mock_pythoncom.CoInitialize.assert_called_once()
        mock_win32com.client.DispatchEx.assert_called_with("Word.Application")
        mock_app.Quit.assert_called_once()
        mock_pythoncom.CoUninitialize.assert_called_once()

    @patch('word_document_master_tool.word.word_app.win32com', None)
    def test_word_app_non_windows(self):
        # Тестируем поведение на не-Windows (или без pywin32)
        with WordApp() as word:
            self.assertIsNone(word.app)
            self.assertFalse(word._owns_app)

if __name__ == '__main__':
    unittest.main()
