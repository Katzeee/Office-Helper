import unittest

class BaseTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        print("===== Starting Test Cases for {} =====".format(cls.__name__))

    @classmethod
    def tearDownClass(cls):
        print("===== Ending Test Cases for {} =====".format(cls.__name__))

    def setUp(self):
        print(f"[ START  ] {self.id()}")

    def tearDown(self):
        # print(f"Finished test: {self._testMethodName}\n")
        pass
