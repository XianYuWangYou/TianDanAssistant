import unittest
import os
import sys
import json
from unittest.mock import patch, MagicMock

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from document_processor import DocumentProcessor
from auto_updater import AutoUpdater


class TestDocumentProcessor(unittest.TestCase):
    """测试文档处理器功能"""

    def setUp(self):
        """测试前准备"""
        self.processor = DocumentProcessor()

    def test_find_placeholders_in_text(self):
        """测试从文本中查找占位符的功能"""
        text = "这是一个测试文本，包含{占位符1}和{占位符2}。"
        expected = {"占位符1", "占位符2"}
        result = self.processor.find_placeholders_in_text(text)
        self.assertEqual(result, expected)

    def test_find_placeholders_in_text_no_placeholders(self):
        """测试没有占位符的文本"""
        text = "这是一个没有占位符的测试文本。"
        expected = set()
        result = self.processor.find_placeholders_in_text(text)
        self.assertEqual(result, expected)


class TestAutoUpdater(unittest.TestCase):
    """测试自动更新功能"""

    def setUp(self):
        """测试前准备"""
        self.updater = AutoUpdater()

    def test_version_initialization(self):
        """测试版本号初始化"""
        self.assertIsNotNone(self.updater.current_version)
        self.assertIsInstance(self.updater.current_version, str)

    @patch('auto_updater.requests.get')
    def test_get_latest_release_success(self, mock_get):
        """测试成功获取最新版本信息"""
        # 模拟API响应
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "tag_name": "v1.2.0",
            "name": "Release 1.2.0",
            "body": "更新内容"
        }
        mock_response.raise_for_status.return_value = None
        mock_get.return_value = mock_response

        result = self.updater.get_latest_release()
        self.assertIsNotNone(result)
        self.assertEqual(result["tag_name"], "v1.2.0")

    @patch('auto_updater.requests.get')
    def test_get_latest_release_failure(self, mock_get):
        """测试获取最新版本信息失败"""
        from requests import RequestException
        mock_get.side_effect = RequestException("网络错误")

        result = self.updater.get_latest_release()
        self.assertIsNone(result)


class TestConfigurationFiles(unittest.TestCase):
    """测试配置文件"""

    def test_app_data_json_exists(self):
        """测试app_data.json文件存在"""
        config_path = os.path.join(os.path.dirname(__file__), 'app_data.json')
        self.assertTrue(os.path.exists(config_path))

    def test_app_data_json_valid(self):
        """测试app_data.json是有效的JSON格式"""
        config_path = os.path.join(os.path.dirname(__file__), 'app_data.json')
        with open(config_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        self.assertIsInstance(data, dict)


def run_all_tests():
    """运行所有测试"""
    # 创建测试加载器
    loader = unittest.TestLoader()
    
    # 创建测试套件
    test_suite = unittest.TestSuite()
    
    # 添加测试用例
    test_suite.addTests(loader.loadTestsFromTestCase(TestDocumentProcessor))
    test_suite.addTests(loader.loadTestsFromTestCase(TestAutoUpdater))
    test_suite.addTests(loader.loadTestsFromTestCase(TestConfigurationFiles))
    
    # 创建测试运行器
    runner = unittest.TextTestRunner(verbosity=2)
    
    # 运行测试
    result = runner.run(test_suite)
    
    # 返回测试结果
    return result.wasSuccessful()


if __name__ == '__main__':
    print("开始运行自动化测试...")
    print("=" * 50)
    
    success = run_all_tests()
    
    print("=" * 50)
    if success:
        print("所有测试通过! ✅")
        sys.exit(0)
    else:
        print("部分测试失败! ❌")
        sys.exit(1)