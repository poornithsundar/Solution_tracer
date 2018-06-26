import pandas as pd
import time

# filenames
excel_names = ["SampleOutput.xlsx", "SampleOutput1.xlsx", "SampleOutput2.xlsx"]

# read them in
excels = [pd.ExcelFile(name) for name in excel_names]

# turn them into dataframes
frames = [x.parse(x.sheet_names[0], header=None,index_col=None) for x in excels]

# concatenate them..
combined = pd.concat(frames)

# write it out
combined.to_excel("Output.xlsx", header=False, index=False)
print('3 excel files concatenated')
time.sleep(2)
