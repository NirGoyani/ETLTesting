import unittest
from HtmlTestRunner import HTMLTestRunner

# Import test cases from both files
from FixedTermDeal_test import FixedTermDealTestCases
from DailyInterestRateForFTD_test import DailyInterestRateForFTDTestCases

def suite():
    test_suite = unittest.TestSuite()
    loader = unittest.TestLoader()
    test_suite.addTests(loader.loadTestsFromTestCase(FixedTermDealTestCases))
    test_suite.addTests(loader.loadTestsFromTestCase(DailyInterestRateForFTDTestCases))
    return test_suite


if __name__ == '__main__':
    runner = HTMLTestRunner(output="./test_reports")
    runner.run(suite())
