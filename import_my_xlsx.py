import pandas as pd


def ImportMyXlsx():
    path = 'sample.xlsx'
    df = pd.read_excel(path, index_col=0)
    val = input('Enter your number: ')

    count = 0
    for item in df.index:
        if item == val:
            num = count
        count += 1

    df_header_index = pd.read_excel(path, header=None, index_col=None)
    my_list = df_header_index.loc[[1, num, num+1]]
    my_list = my_list.values.tolist()
    y_m = YearMonth(df_header_index)

    tmp = 0
    for day in my_list[0]:
        try:
            start
            # 変数startが存在する場合は以降の処理が実行
            if type(day) == str:
                end = tmp
                break
        except NameError:
            # 変数startが存在しない場合は以降の処理が実行
            if day == 1:
                start = tmp
        tmp += 1

    list_day = my_list[0][start:end]
    list_start = my_list[1][start:end]
    list_end = my_list[2][start:end]

    result_list = []
    koukyu = 0
    yukyu = 0

    for index in range(len(list_day)):
        if list_start[index] == "●":
            # print('公休')
            koukyu += 1
        elif list_start[index] == "有":
            # print('有休')
            yukyu += 1
        else:
            if list_day[index] < 10:
                list_day[index] = "0" + str(list_day[index])
            else:
                list_day[index] = str(list_day[index])
            result_list.append([list_day[index], list_start[index], list_end[index]])
    print("公給：" + str(koukyu) + "日")
    print("有給：" + str(yukyu) + "日")
    print("休日：" + str(koukyu + yukyu) + "日")
    print("出勤：" + str(len(result_list)) + "日")

    return [y_m, result_list]


def YearMonth(title):
    tmp = title.loc[0][0]
    y = tmp.find('年')
    m = y + 1
    y = y - 4
    # print(tmp[y:y+4])
    # print(tmp[m:m+2])
    y_m = [tmp[y:y+4], tmp[m:m+2]]

    return y_m