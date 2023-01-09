import pandas as pd
from yt_stats import YTstats
from tqdm import tqdm
from datetime import datetime



now = datetime.now()

current = now.strftime("%H%M")


API_KEY = 'AIzaSyCNxDKfMX8TJ8DBA2fFpN-HXGJzs4VZghc'

df = pd.read_excel("HereIAm.xlsx")
Views =[]
Comments = []

for x,y in tqdm(zip(df["Video ID"],df["Channel ID"])):
    channel_id = y
    video_id = x
    part = 'statistics'
    yt = YTstats(API_KEY, channel_id)
    a = yt._get_single_video_data(video_id,part)
    Views.append(a['viewCount'])
    Comments.append(a['commentCount'])
    

df["views"] = Views
df["Comments"] = Comments
del df["Channel ID"]
del df["Video ID"]



col = len(df.columns)

df.sort_values(['views'], ascending=False, inplace=True)
df.reset_index(inplace=True)
df.index += 1



# def highlight(row):
#     pink = 'background-color: #FFC0CB'
#     blue = 'background-color: #0000FF'
#     print(row['Group'])
#     if row['Group'] > 1:
#         print("hey")
#         return ['background-color: #FFC0CB']
#     else:
#         return [blue]

# writer = pd.ExcelWriter(f'{current}.xlsx', engine='xlsxwriter')

with pd.ExcelWriter(rf'C:/Users/Annie/Desktop/Coding/python/youtube data project/{current}.xlsx', engine='xlsxwriter') as writer:
    del df["index"]
    df.to_excel(writer, sheet_name='Sheet1')

# style.apply(highlight, axis=1)


# df.to_excel(writer, f"{current}.xlsx", sheet_name="Sheet1")

# workbook  = writer.book
# worksheet = writer.sheets['Sheet1']
# pink = workbook.add_format({'bg_color': '#FFC0CB'})
# blue = workbook.add_format({'bg_color': '#0000FF'})


# for i in range(0, len(df)):
#     df['Group'].values[i]
#     if df['Group'].values[i] == "G":
#         print("Hello")
#         worksheet.conditional_format(i, 0, i, col, 
#         {'type': 'no_blanks', 'format':   pink})
#     else:
#         worksheet.conditional_format(i, 0, i, col, {'type': 'no_blanks', 
#         'format':   blue})



