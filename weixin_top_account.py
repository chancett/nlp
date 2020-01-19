import glob, os
import pandas as pd

def top_view_author():
    df = pd.DataFrame()
    for file in glob.glob(os.path.join(r'S:\新邦\2019M12nike_word', '*.json')):
        print(file)
        temp = pd.read_json(file, lines=True, encoding='utf-8')
        temp = temp.drop_duplicates(subset = ['author', 'uuid'])
        temp = temp[['author', 'author_id', 'uuid', 'views', 'likes']]
        df = df.append(temp)
    df = df.drop_duplicates(subset = ['author', 'uuid'])
    result = df.pivot_table(index=['author', 'author_id'], aggfunc={'views': 'sum'}).reset_index()
    result = result.sort_values(by='views', ascending=False)
    result.to_excel(r'R:\yuqing_fmcg\3_report\2_Monthly_report\201912\微信浏览量top作者所发文章热词\top_author.xlsx', index=False)

dec_top_author = ['flightclub', 'qqnba-wx', 'supreme007com', 'youzzu', 'XCin666',
       'poizonapp', 'girlnba', 'snkrmania', 'lol_helper', 'lanqiujiaoxue',
       'BallTRK', 'kuangyandoggy', 'ilianyujia', 'yangyitalk', 'lqyodu',
       'gossipleague', 'lanqiujueji', 'quanshangcn', 'UNIQLO_CHINA',
       'dddnba', 'Yoga-Road', 'instachina', 'SINA_NBA', 'zqcf518',
       'ckoome', 'ka24025', 'suqunbasketball', 'swagdog', 'sxgqtwx',
       'HUPUtiyu']

def top_veiw_article():
    df = pd.DataFrame()
    for file in glob.glob(os.path.join(r'S:\新邦\2019M12nike_word', '*.json')):
        print(file)
        temp = pd.read_json(file, lines=True, encoding='utf-8')
        temp = temp.drop_duplicates(subset=['author_id', 'uuid'])
        temp = temp[temp['author_id'].isin(dec_top_author)]
        df = df.append(temp)
    df = df.drop_duplicates(subset=['author_id', 'uuid'])
    writer = pd.ExcelWriter(r'R:\yuqing_fmcg\3_report\2_Monthly_report\201912\top_author_article.xlsx', options={'strings_to_urls': False})
    df.to_excel(writer, index=False)
    writer.save()

if __name__ == '__main__':
    top_veiw_article()