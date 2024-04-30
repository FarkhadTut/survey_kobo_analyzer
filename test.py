import pandas as pd

# Sample DataFrame
data = {'A': ['foo', 'foo', 'foo', 'foo', 'bar', 'bar', 'bar', 'bar'],
        'B': ['one', 'one', 'two', 'two', 'one', 'one', 'two', 'two'],
        'C': ['small', 'large', 'large', 'small', 'small', 'large', 'small', 'small']}
df = pd.DataFrame(data)

# Count occurrences of each combination of A and B
cross_tab = pd.crosstab(df['A'], df['B'])
print(cross_tab)