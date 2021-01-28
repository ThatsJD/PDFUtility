import unittest
import main
class MyTestCase(unittest.TestCase):
    def test_something(self):
        val = main.Docx2PDFConvert('D:/PycharmProjects/PDFBasicUtilities/file.docx', 'D:/PycharmProjects/PDFBasicUtilities/newfile.pdf')
        print(val)
        self.assertEqual(True, False)


if __name__ == '__main__':
    unittest.main()
