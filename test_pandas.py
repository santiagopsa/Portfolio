import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


df = pd.read_excel(r'statements/bancolombia_clean.xlsx')
print(df.columns)
updated_df=df.drop(["FECHA","VALOR","SUCURSAL"], axis=1).head()
print(updated_df.head())
