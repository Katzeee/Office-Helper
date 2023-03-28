import unittest
import os

class CustomTestResult(unittest.TextTestResult):
    def addSuccess(self, test):
        super().addSuccess(test)
        print(f"[   OK   ] {test.id()}")

    def addError(self, test, err):
        super().addError(test, err)
        print(f"[ ERROR  ] {test.id()}")

    def addFailure(self, test, err):
        super().addFailure(test, err)
        print(f"[ FAILED ] {test.id()}")

if __name__ == '__main__':
    os.chdir(os.path.abspath(os.path.dirname(__file__)))
    # discover all test cases in the 'tests' directory
    test_suite = unittest.defaultTestLoader.discover(start_dir='tests')

    # create a test runner
    runner = unittest.TextTestRunner(resultclass=CustomTestResult)

    # run the test suite
    runner.run(test_suite)
