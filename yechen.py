import pandas as pd
from yt_stats import YTstats
from tqdm import tqdm
from datetime import datetime



now = datetime.now()

current = now.strftime("%H%M")


API_KEY = 'AIzaSyCNxDKfMX8TJ8DBA2fFpN-HXGJzs4VZghc'

df = pd.read_excel("yechen.xlsx")
Views =[]
Comments = []
# Likes = []

for x,y in tqdm(zip(df["Video ID"],df["Channel ID"])):
    channel_id = y
    video_id = x
    part = 'statistics'
    yt = YTstats(API_KEY, channel_id)
    a = yt._get_single_video_data(video_id,part)
    Views.append(a['viewCount'])
    Comments.append(a['commentCount'])
    # if a['likeCount']:
    #     Likes.append(a['likeCount'])

df["views"] = Views
df["Comments"] = Comments
# df["like"] = Likes
del df["Channel ID"]
del df["Video ID"]


with pd.ExcelWriter(rf'C:/Users/Annie/Desktop/Coding/python/youtube data project/{current}yc.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False)




