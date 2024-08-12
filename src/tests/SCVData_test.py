import sys
import os
import glob
import unittest
import logging
import pandas as pd

current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(current_dir, '..', '..'))
sys.path.append(project_root)

from dotenv import load_dotenv
dotenv_path = os.path.join(project_root, '.env')
load_dotenv(dotenv_path)

from src.utility.DownloadTheData import DownloadData

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class SCVDataTestCases(unittest.TestCase):

    @staticmethod
    def find_and_read_csv(directory, pattern):
        # Construct the search pattern
        search_pattern = os.path.join(directory, pattern)
        # Find the files matching the pattern
        files = glob.glob(search_pattern)

        if not files:
            logger.error(f"No files found matching pattern {pattern} in directory {directory}")
            return None

        # Read the first matching file
        file_to_read = files[0]
        logger.info(f"Reading file: {file_to_read}")

        # Read the CSV file into a DataFrame
        df = pd.read_csv(file_to_read, delimiter="|")

        return df

    @classmethod
    def setUpClass(cls):
        logger.info("Data is Downloading....")

        expected_data_path = os.path.join(current_dir, '..', 'tempStorage', 'expectedData')

        cls.expected_data_dir = os.path.normpath(expected_data_path)

        downloader = DownloadData()
        downloader.get_files_modified_on_latest_date("DEV/SCV", cls.expected_data_dir)

        # Read the CSV data
        cls.df_contactDetails = cls.find_and_read_csv(cls.expected_data_dir, "*Contactdetails.csv")
        cls.df_customerDetails = cls.find_and_read_csv(cls.expected_data_dir, "*Customerdetails.csv")
        cls.df_detailsOfAccount = cls.find_and_read_csv(cls.expected_data_dir, "*Detailsofaccount.csv")
        cls.df_AggregateBalanceDetails = cls.find_and_read_csv(cls.expected_data_dir, "*Aggregatebalancedetails.csv")

        logger.info("Data Downloaded")

    def test_to_verify_contact_details_from_SCV_data(self):
        addressLine1 = self.df_contactDetails['Address Line 1'].tolist()
        addressLine2 = self.df_contactDetails['Address Line 2'].tolist()
        post_code = self.df_contactDetails['Post Code'].tolist()

        # This three column should have some data
        for idx in range(len(addressLine1) - 1):
            line1 = addressLine1[idx]
            line2 = addressLine2[idx]
            postCode = post_code[idx]

            self.assertTrue(isinstance(line1, str) and line1.strip(),
                            f"Address Line 1 is empty at row {idx + 1}")
            self.assertTrue(isinstance(line2, str) and line2.strip(),
                            f"Address Line 2 is empty at row {idx + 1}")
            self.assertTrue(isinstance(postCode, str) and postCode.strip(),
                            f"Postcode is empty at row {idx + 1}")


    def test_to_verify_customer_details_from_SCV_data(self):
        companyName = self.df_customerDetails['Customer Surname Or Company Name'].tolist()
        companyNumber = self.df_customerDetails['Company Number'].tolist()

        # This loop should check if the specific values exist
        found_comName = False
        found_comNumber = False

        for idx in range(len(companyName)-1):
            comName = companyName[idx]
            comNumber = companyNumber[idx]

            if comName == "Flagstone Investment Management Limited":
                found_comName = True
            if comNumber == 08528880.0:
                found_comNumber = True

        self.assertTrue(found_comName, "The company name 'Flagstone Investment Management Limited' was not found.")
        self.assertTrue(found_comNumber, "The company number '08528880' was not found.")

    def test_to_verify_details_of_accounts_from_SCV_data(self):
        accountTitle = self.df_detailsOfAccount['Account Title'].tolist()
        accountNumber = self.df_detailsOfAccount['Account Number'].tolist()
        productName = self.df_detailsOfAccount['Product Name'].tolist()
        accountHolderIndicator = self.df_detailsOfAccount['Account Holder Indicator'].tolist()
        accountStatusCode = self.df_detailsOfAccount['Account Status Code'].tolist()
        exclusionType =  self.df_detailsOfAccount['Exclusion Type'].tolist()
        recentTransactions = self.df_detailsOfAccount['Recent Transactions'].tolist()
        accountBranchJurisdiction = self.df_detailsOfAccount['Account Branch Jurisdiction'].tolist()
        BRRDMarking = self.df_detailsOfAccount['BRRD Marking'].tolist()
        structuredDepositAccounts = self.df_detailsOfAccount['Structured Deposit Accounts'].tolist()
        accountBalanceInSterling = self.df_detailsOfAccount['Account Balance in Sterling'].tolist()
        authorisedNegativeBalance = self.df_detailsOfAccount['Authorised Negative Balances'].tolist()
        currencyOfAccount = self.df_detailsOfAccount['Currency of Account'].tolist()
        accountBalanceInOriginalCurrency = self.df_detailsOfAccount['Account Balance in Original Currency'].tolist()
        exchangeRate = self.df_detailsOfAccount['Exchange Rate'].tolist()
        originalAccountBalanceBeforeInterest = self.df_detailsOfAccount['Original Account Balance Before Interest'].tolist()

        found_accountTitle = False
        found_accountNumber = False
        found_productName = False
        found_accountHolderIndicator = False
        found_accountStatusCode = False
        found_exclusionType = False
        found_recentTransactions = False
        found_accountBranchJurisdiction = False
        found_BRRDMarking = False
        found_structuredDepositAccounts = False
        found_accountBalanceInSterling = False
        found_authorisedNegativeBalance = False
        found_currencyOfAccount = False
        found_accountBalanceInOriginalCurrency = False
        found_exchangeRate = False
        found_originalAccountBalanceBeforeInterest = False

        logger.info(f"accountNumber values: {accountNumber}")

        # This loop should check if the specific values exist
        for idx in range(len(accountTitle) - 1):

            if accountTitle[idx] == "Flagstone Group LTD Client Account":
                found_accountTitle = True
            if accountNumber[idx] == "SILFIM01YSHAHEEN":
                found_accountNumber = True
            if productName[idx] == "Fixed Term Savings":
                found_productName = True
            if accountHolderIndicator[idx] == 1:
                found_accountHolderIndicator = True
            if accountStatusCode[idx] == "B":
                found_accountStatusCode = True
            if exclusionType[idx] == "BEN":
                found_exclusionType = True
            if recentTransactions[idx] == "Yes":
                found_recentTransactions = True
            if accountBranchJurisdiction[idx] == "GBR":
                found_accountBranchJurisdiction = True
            if BRRDMarking[idx] == "Yes":
                found_BRRDMarking = True
            if structuredDepositAccounts[idx] == "No":
                found_structuredDepositAccounts = True
            if accountBalanceInSterling[idx] == 100000.00:
                found_accountBalanceInSterling = True
            if authorisedNegativeBalance[idx] == 0:
                found_authorisedNegativeBalance = True
            if currencyOfAccount[idx] == "GBP":
                found_currencyOfAccount = True
            if accountBalanceInOriginalCurrency[idx] == 100000.00:
                found_accountBalanceInOriginalCurrency = True
            if exchangeRate[idx] == 1.000000000:
                found_exchangeRate = True
            if originalAccountBalanceBeforeInterest[idx] == 100000.00:
                found_originalAccountBalanceBeforeInterest = True

        self.assertTrue(found_accountTitle, "The Account title was not found.")
        self.assertTrue(found_accountNumber, "The Account Number was not found.")
        self.assertTrue(found_productName, "The Product Number was not found.")
        self.assertTrue(found_accountHolderIndicator, "The account Holder indicator was not found.")
        self.assertTrue(found_accountStatusCode , "The account status code was not found")
        self.assertTrue(found_exclusionType , "The exclusion type was not found.")
        self.assertTrue(found_recentTransactions , "The recent transactions was not found")
        self.assertTrue(found_accountBranchJurisdiction , "The account branch jurisdiction was not found.")
        self.assertTrue(found_BRRDMarking, "The BRRD marking was not found.")
        self.assertTrue(found_structuredDepositAccounts , "The structure deposit accounts was not found")
        self.assertTrue(found_accountBalanceInSterling , "The account balance in sterling was not found")
        self.assertTrue(found_authorisedNegativeBalance , "The autorise negative balance was not found")
        self.assertTrue(found_currencyOfAccount , "the currency of account was not found")
        self.assertTrue(found_accountBalanceInOriginalCurrency , "The account balance in original currency was not found")
        self.assertTrue(found_exchangeRate , "The exchange rate was not found")
        self.assertTrue(found_originalAccountBalanceBeforeInterest , "The original account balance before interest")
        self.assertTrue(True)

    def test_to_verify_aggregated_balance_details_from_SCV_data(self):
        aggregateBalance = self.df_AggregateBalanceDetails['Aggregate Balance'].tolist()
        compensatableAmount = self.df_AggregateBalanceDetails['Compensatable Amount'].tolist()

        found_aggregateBalance = False
        found_compensatableAmount = False

        for idx in range(len(aggregateBalance) - 1):
            if aggregateBalance[idx] == 13802905.18:
                found_aggregateBalance = True
            if compensatableAmount[idx] == 13802905.18:
                found_compensatableAmount = True

        self.assertTrue(found_aggregateBalance , "The aggregate balance was not found")
        self.assertTrue(found_compensatableAmount , "The compensate amount was not found")

        self.assertTrue(True)

    def test_to_verify_SCVID_is_present_in_all_other_SCV_dataset(self):
        # Extract all values from the SCVID column
        if self.df_contactDetails is not None:
            scvid_values = self.df_contactDetails['SCVID'].tolist()
            logger.info(f"SCVID values: {scvid_values}")

            # Check if SCVIDs are present in all other DataFrames
            if self.df_customerDetails is not None:
                self.assertTrue(
                    self.df_customerDetails['SCVID'].isin(scvid_values).all(),
                    "Not all SCVIDs in df_contactDetails are present in df_customerDetails"
                )
            if self.df_detailsOfAccount is not None:
                self.assertTrue(
                    self.df_detailsOfAccount['SCVID'].isin(scvid_values).all(),
                    "Not all SCVIDs in df_contactDetails are present in df_detailsOfAccount"
                )
            if self.df_AggregateBalanceDetails is not None:
                self.assertTrue(
                    self.df_AggregateBalanceDetails['SCVID'].isin(scvid_values).all(),
                    "Not all SCVIDs in df_contactDetails are present in df_AggregateBalanceDetails"
                )

    @classmethod
    def tearDownClass(cls):
        logger.info("Tearing down resources")
        cls.clean_up_downloaded_files(cls.expected_data_dir)

    @classmethod
    def clean_up_downloaded_files(cls, folder_path):
        logger.info(f"Cleaning up files in {folder_path}...")
        files = glob.glob(os.path.join(folder_path, "*"))
        for file in files:
            os.remove(file)
            logger.info(f"Deleted file: {file}")
        logger.info("Clean up completed.")


if __name__ == '__main__':
    unittest.main()
