import pandas as pd
df = pd.DataFrame({'a': ['1\n', '2\n', '3'], 'b': ['4\n', '5', '6\n']})

df.replace({'\n': '<br>'}, regex=True, inplace=True)

print(df)