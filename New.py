import pandas as pd
df = pd.read_csv(r"C:\Users\gdkvarma\Documents\My Projects\Time_Calculations-master\Timedata.csv",usecols=['Event Date','Location','Card Number'],dtype=object)
df['Event Date'] = pd.to_datetime(df['Event Date'])
df['IN/OUT']= [io.split(" ")[-1] for io in df['Location']]
df.sort_values(['Card Number', 'Event Date'], inplace=True, axis=0)
df.drop(['Location'],inplace=True, axis=1)
df['Date'] = df['Event Date'].dt.date
df_outtime, df_intime = [x for _, x in df.groupby(df['IN/OUT'] == "IN")]
i,o = df_intime['Event Date'],df_outtime['Event Date']
df['In'],df['Out'] = i,o
df.drop(['IN/OUT','Event Date'],inplace=True, axis=1)
df['In'].fillna(method='ffill', inplace=True)
df['Out'].fillna(method='backfill', inplace=True)
df.drop_duplicates(keep='first', inplace=True)
df['Working Hours'] = df['Out'] -df ['In']
df['In'],df['Out'] = df['In'].dt.time,df['Out'].dt.time
d = df.groupby(["Card Number","Date"],as_index=False).sum()
d['Working Hours'] = [w.split(" ")[-1] for w in d['Working Hours'].astype(str)]
d['Date'] = [w.split(" ")[0] for w in d['Date'].astype(str)]
e = d.pivot('Card Number', 'Date').rename_axis().fillna('A')
e.to_csv("by_date.csv")
df['Working Hours'] = [w.split(" ")[-1] for w in df['Working Hours'].astype(str)]
