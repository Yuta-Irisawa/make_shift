import openpyxl
import pandas as pd
import glob

def input_from_excel(foldername):
    filespath = foldername + "\*"
    files = glob.glob(filespath)

    df = pd.DataFrame()

    for i,file2 in enumerate(files):
        if i == len(files)-2:
            break
        df2 = pd.read_excel(file2)
        name2 = df2.columns[2]
        df2 = df2.rename(columns={'Unnamed: 0':'日付', '名前　→':'曜日', 'Unnamed: 3':'午前', 'Unnamed: 4':'午後', '{}'.format(name2):'必要人数'})
        df2 = df2[3:]
        df2 = df2.drop('曜日', axis=1)
        df2['名前'] = name2
        df2['回数'] = 0
        dropindex2 = df2.index[(df2['午前']=="x") & (df2['午後']=="x")]
        df2 = df2.drop(dropindex2)
        
        df = pd.concat([df, df2])

    df = df.sort_values('日付')
    df = df.reset_index()
    return df


def organize_df(df):
    date_member_num_dict = {}
    date_necessary = {}
    date_priority = {}
    date_priority_half = {}

    for date_num in range(len(df['日付'].unique())):
        date = df['日付'].unique()[date_num]
        date_member_num = 0
        date_am_num, date_pm_num = 0, 0
        for i, d_df in enumerate(df['日付']):
            necessary_num = 0
            if d_df == date:
                necessary_num = df['必要人数'][i]
                while i < len(df) and d_df == df['日付'][i]:
                    if df['午前'][i] != "x":
                        date_am_num += 1
                    if df['午後'][i] != "x":
                        date_pm_num += 1
                    date_member_num += 1
                    i += 1
                break
        date_member_num_dict['{}'.format(date)] = date_member_num
        date_priority['{}'.format(date)] = date_member_num - necessary_num
        date_necessary['{}_am'.format(date)] = necessary_num
        date_necessary['{}_pm'.format(date)] = necessary_num
        date_priority_half['{}_am'.format(date)] = date_am_num - necessary_num
        date_priority_half['{}_pm'.format(date)] = date_pm_num - necessary_num
        
    date_priority = sorted(date_priority.items(), key=lambda x:x[1])
    date_priority_half = sorted(date_priority_half.items(), key=lambda x:x[1])

    return date_necessary, date_priority_half

def make_workmember_list(date_priority_half, date_necessary, df):
    work_member_list = []
    for i in range(len(date_priority_half)):
        date_info = {}
        priority = date_priority_half[i][0]
        date_info['日付'] = priority
        date_info['必要人数'] = date_necessary[priority]
        remain = date_necessary[priority]
        
        df = df.sort_values('回数')
        df = df.reset_index(drop=True)

        for j,d_df in enumerate(df['日付']):
            member = df['名前'][j]
            am = df['午前'][j]
            pm = df['午後'][j]
            if str(d_df) in priority and "am" in priority and remain > 0 and am != "x":
                date_info['{}'.format(member)] = "o"
                remain -= 1
                for k,m_df in enumerate(df['名前']):
                    if m_df == member:
                        df_num = int(df['回数'][k])
                        df['回数'][k] = df_num + 1
            elif str(d_df) in priority and "pm" in priority and remain > 0 and pm != "x":
                date_info['{}'.format(member)] = "o"
                remain -= 1
                for k,m_df in enumerate(df['名前']):
                    if m_df == member:
                        df_num = int(df['回数'][k])
                        df['回数'][k] = df_num + 1
        work_member_list.append(date_info)

    return work_member_list

def work_member_list2excel(work_member_list, foldername):
    a_df = pd.DataFrame(work_member_list)
    a_df_sorted = a_df.sort_values(['日付'])
    a_df_sorted.to_excel('{}_shift.xlsx'.format(foldername), sheet_name='sheet1', index=False)

def main():
    foldername = input("入力ExcelファイルがまとまっているフォルダのPathを入力してください。")
    df = input_from_excel(foldername)
    date_necessary, date_priority_half = organize_df(df)
    work_member_list = make_workmember_list(date_priority_half, date_necessary, df)
    work_member_list2excel(work_member_list, foldername)

if __name__ == '__main__':
    main()