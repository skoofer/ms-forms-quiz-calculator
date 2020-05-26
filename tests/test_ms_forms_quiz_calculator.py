import os
import unittest
import pandas as pd

from ms_forms_quiz_calculator.ms_forms_quiz_calculator import calculate_quiz_results


class MSFormsQuizCalculator(unittest.TestCase):
    def test_results(self):
        calculate_quiz_results('resources/quizzes', 'actual_results')

        expected_file = 'resources/expected_results.xlsx'
        df_expected = self.read_excel_into_df(expected_file)

        expected_file = 'resources/quizzes/actual_results.xlsx'
        df_actual = self.read_excel_into_df(expected_file)

        os.remove('resources/quizzes/actual_results.xlsx')

        self.assertTrue(df_expected.equals(df_actual))

    @staticmethod
    def read_excel_into_df(file_name):
        file = pd.ExcelFile(file_name)
        df = pd.DataFrame()
        df = df.append(pd.read_excel(file, 'Category-1'))
        df = df.append(pd.read_excel(file, 'Category-2'))
        df = df.append(pd.read_excel(file, 'Category-3'))
        df = df.append(pd.read_excel(file, 'Grand Total'))

        return df


if __name__ == '__main__':
    unittest.main()
