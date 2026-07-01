import pandas as pd
df = pd.read_csv(r'C:\Users\Researcher\Desktop\Project V\OCR Sample\discount_rate_pipeline\peer_group_features_enriched_partial.csv', encoding='utf-8-sig')
print(df['track_type'].value_counts())
print('tech_special:', int((df['is_tech_special']==1).sum()))
print('general & not tech:', int(((df['track_type']=='general') & (df['is_tech_special']==0)).sum()))
