import pandas as pd
import pytest
import xlsxwriter
from helpers.utils import load_excel_sheets  



def test_load_excel_sheets():
    # Setup: Create a sample Excel file with known data
    file_path = 'test_data.xlsx'
    df1 = pd.DataFrame({'A': [1, 2, 3]})
    df2 = pd.DataFrame({'B': [4, 5, 6]})
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df1.to_excel(writer, sheet_name='Sheet1', index=False)
        df2.to_excel(writer, sheet_name='Sheet2', index=False)

    # Call the function with the test file and sheet names
    result = load_excel_sheets(file_path, ['Sheet1', 'Sheet2'])

    # Correctly reset the index of the DataFrames read from the file for comparison
    # while maintaining the tuple type
    result = tuple(df.reset_index(drop=True) for df in result)

    # Assertions to check if the function works as expected
    assert isinstance(result, tuple), "Result should be a tuple"
    assert len(result) == 2, "Tuple should have two elements"
    assert all(isinstance(df, pd.DataFrame) for df in result), "All tuple elements should be DataFrames"
    assert result[0].equals(df1), "First DataFrame should match df1"
    assert result[1].equals(df2), "Second DataFrame should match df2"

    # Clean up: Remove the test file
    import os
    os.remove(file_path)
