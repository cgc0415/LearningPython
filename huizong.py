import xlrd
import xlwt

wb = xlrd.open_workbook("D:\\Programming\\Python\\data\\苍溪1月.xls")   #打开名为苍溪的Excel

sheetnames = wb.sheet_names()   #获取苍溪的所有表名
for sht_nam in sheetnames:
    if (sht_nam.find('附件') != -1 or sht_nam == 'Worksheet'):  # find-在字符串中查找某子串，找不到则返回-1
        a = 1
    else:
        dstfile = xlwt.Workbook()
        newtalbe = dstfile.add_sheet('info', cell_overwrite_ok=True)
        sheet = wb.sheet_by_name(sht_nam)  # 第一个表
        destFullName = 'D:\\Programming\\Python\\data\\' + sht_nam + '_output.xls'
        print(sht_nam)

        # 数据正文起始行索引
        g_row_data = 6

        g_xiaoliang = 0  # 销量 AE+AD

        g_col_dangci = 2  # 档次所在列
        g_col_guige = 9  # 规格所在列
        g_col_kaixiang = 10  # 开向所在列
        g_col_huase = 12  # 花色所在列
        g_col_zhizaobumen = 25

        # 当日库房结存，列索引：30.
        g_col_kufangxiaoshou = 28
        g_col_kufangqita = 29
        g_col_kufangjiecun = 30
        g_youxiao_kucun = sheet.row_values(4)[g_col_kufangjiecun]
        g_wuxiao_kucun = sheet.row_values(5)[g_col_kufangjiecun]

        # 获取表sheet的总行数、列数
        nrows = sheet.nrows
        ncols = sheet.ncols
        print(nrows, ncols)

        # 创建一个list用于保存表的数据
        list = [[0] * ncols for i in range(nrows)]

        # 将表的数据赋值给list
        for i in range(0, nrows):
            for j in range(0, ncols):
                list[i][j] = sheet.row_values(i)[j]

        # 统计锁的结存数
        g_lock_record_num = 0  # 带锁的记录条数
        g_lock_num = 0  # 定义该变量用于统计所有锁的结存数量
        g_lock_xiaoliang = 0  # 锁的销量
        for i in range(0, nrows):
            if (list[i][2].find('锁') != -1):  # find-在字符串中查找某子串，找不到则返回-1
                g_lock_record_num += 1
                g_lock_num += list[i][g_col_kufangjiecun]
                g_lock_xiaoliang = g_lock_xiaoliang + sheet.row_values(i)[g_col_kufangxiaoshou] + sheet.row_values(i)[g_col_kufangqita]

        # 有效库存数要减去锁的数量
        g_youxiao_kucun -= g_lock_num

        # 写入excel的行号
        g_write_row_index = 0
        newtalbe.write(g_write_row_index, 0, '名称')
        newtalbe.write(g_write_row_index, 1, '记录数')
        newtalbe.write(g_write_row_index, 2, '销量')
        newtalbe.write(g_write_row_index, 3, '结存')
        g_write_row_index += 1
        print('开向', '记录数', '销量', '结存')

        newtalbe.write(g_write_row_index, 0, '智能锁')
        newtalbe.write(g_write_row_index, 1, g_lock_record_num)
        newtalbe.write(g_write_row_index, 2, g_lock_xiaoliang)
        newtalbe.write(g_write_row_index, 3, g_lock_num)
        g_write_row_index += 1

        newtalbe.write(g_write_row_index, 0, '无效')
        newtalbe.write(g_write_row_index, 1, 1)
        newtalbe.write(g_write_row_index, 2,sheet.row_values(5)[g_col_kufangxiaoshou] + sheet.row_values(5)[g_col_kufangqita])
        newtalbe.write(g_write_row_index, 3, sheet.row_values(5)[g_col_kufangjiecun])
        g_write_row_index += 1

        newtalbe.write(g_write_row_index, 0, '有效')
        newtalbe.write(g_write_row_index, 1, 1)
        newtalbe.write(g_write_row_index, 2, sheet.row_values(4)[g_col_kufangxiaoshou] + sheet.row_values(4)[g_col_kufangqita] - g_lock_xiaoliang)
        newtalbe.write(g_write_row_index, 3, g_youxiao_kucun)
        g_write_row_index += 1

        print('汇总', '记录数', '销量', '结存')
        print('智能锁', g_lock_record_num, g_lock_xiaoliang, g_lock_num)
        print('无效', '1', sheet.row_values(5)[g_col_kufangxiaoshou] + sheet.row_values(5)[g_col_kufangqita],
              sheet.row_values(5)[g_col_kufangjiecun])
        print('有效', '1',
              sheet.row_values(4)[g_col_kufangxiaoshou] + sheet.row_values(4)[g_col_kufangqita] - g_lock_xiaoliang,
              g_youxiao_kucun)

        print('\n**********开向统计-开始**********')
        g_write_row_index += 1
        # 获取开向种类数
        list_kaixiang = [0] * (nrows - g_row_data)

        for i in range(g_row_data, nrows):
            if (sheet.row_values(i)[2].find('锁') == -1):  # find-在字符串中查找某子串，找不到则返回-1，此处只统计非锁的项目
                list_kaixiang[i - g_row_data] = sheet.row_values(i)[g_col_kaixiang]

        list_kaixiang_count = set(list_kaixiang)
        num_kaixiang = len(list_kaixiang_count)

        list_statics = [[0] * 4 for i in range(num_kaixiang)]  # 4列分别为：开向、数据条数、销量、结存
        i = 0
        for name in list_kaixiang_count:
            list_statics[i][0] = name
            i += 1

        # 处理原始数据，统计每种开向的销量和结存信息
        for i in range(g_row_data, nrows):
            for j in range(0, num_kaixiang):
                if (list_statics[j][0] == sheet.row_values(i)[g_col_kaixiang]):
                    list_statics[j][1] += 1
                    print('error:',i,j,list_statics[j][0], list_statics[j][2], sheet.row_values(i)[g_col_kufangxiaoshou], sheet.row_values(i)[g_col_kufangqita])
                    list_statics[j][2] = list_statics[j][2] + sheet.row_values(i)[g_col_kufangxiaoshou] + sheet.row_values(i)[g_col_kufangqita]
                    list_statics[j][3] += sheet.row_values(i)[g_col_kufangjiecun]

        print('开向', '记录数', '销量', '结存')
        for i in range(0, num_kaixiang):
            if (list_statics[i][0] == '内左'):
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                print(list_statics[i])
                break
            elif (i == num_kaixiang - 1 and list_statics[i][0] != '内左'):
                newtalbe.write(g_write_row_index, 0, '内左')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)
        g_write_row_index += 1

        for i in range(0, num_kaixiang):
            if (list_statics[i][0] == '内右'):
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                print(list_statics[i])
                break
            elif (i == num_kaixiang - 1 and list_statics[i][0] != '内右'):
                newtalbe.write(g_write_row_index, 0, '内右')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)
        g_write_row_index += 1

        for i in range(0, num_kaixiang):
            if (list_statics[i][0] == '外左'):
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                print(list_statics[i])
                break
            elif (i == num_kaixiang - 1 and list_statics[i][0] != '外左'):
                newtalbe.write(g_write_row_index, 0, '外左')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)
        g_write_row_index += 1

        for i in range(0, num_kaixiang):
            if (list_statics[i][0] == '外右'):
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                print(list_statics[i])
                break
            elif (i == num_kaixiang - 1 and list_statics[i][0] != '外右'):
                newtalbe.write(g_write_row_index, 0, '外右')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)
        g_write_row_index += 1

        print('**********开向统计-结束**********')

        print('\n**********档次统计-开始**********')
        g_write_row_index += 1
        # 获取档次种类数
        # ：70钢套、90钢套、外购
        list_gangtao = [[0] * 4 for i in range(3)]  # 4列分别为：规格、数据条数、销量、结存

        list_dangci = [0] * (nrows - g_row_data)

        for i in range(g_row_data, nrows):
            if (sheet.row_values(i)[2].find('锁') == -1):  # find-在字符串中查找某子串，找不到则返回-1，此处只统计非锁的项目
                list_dangci[i - g_row_data] = sheet.row_values(i)[g_col_dangci]

        list_dangci_count = set(list_dangci)
        list_dangci_count.add('外购')
        num_dangci = len(list_dangci_count)

        list_statics = [[0] * 4 for i in range(num_dangci)]  # 4列分别为：规格、数据条数、销量、结存
        i = 0
        for name in list_dangci_count:
            list_statics[i][0] = name
            i += 1

        # 处理原始数据，统计每种档次的销量和结存信息
        for i in range(g_row_data, nrows):
            for j in range(0, num_dangci):
                if (list_statics[j][0] == sheet.row_values(i)[g_col_dangci]):
                    list_statics[j][1] += 1
                    list_statics[j][2] = list_statics[j][2] + sheet.row_values(i)[g_col_kufangxiaoshou] + sheet.row_values(i)[g_col_kufangqita]
                    list_statics[j][3] += sheet.row_values(i)[g_col_kufangjiecun]
                    if (sheet.row_values(i)[g_col_zhizaobumen] == '外购'):
                        list_gangtao[2][0]='外购'
                        list_gangtao[2][1] += 1
                        list_gangtao[2][2] = list_gangtao[2][2] + sheet.row_values(i)[g_col_kufangxiaoshou] + sheet.row_values(i)[g_col_kufangqita]
                        list_gangtao[2][3] += sheet.row_values(i)[g_col_kufangjiecun]
                    if (sheet.row_values(i)[g_col_dangci] == '钢套门' and sheet.row_values(i)[g_col_dangci + 1].find('70') != -1):
                        list_gangtao[0][1] += 1
                        list_gangtao[0][2] = list_gangtao[0][2] + sheet.row_values(i)[g_col_kufangxiaoshou] + sheet.row_values(i)[g_col_kufangqita]
                        list_gangtao[0][3] += sheet.row_values(i)[g_col_kufangjiecun]
                    elif (sheet.row_values(i)[g_col_dangci] == '钢套门' and sheet.row_values(i)[g_col_dangci + 1].find('90') != -1):
                        list_gangtao[1][1] += 1
                        list_gangtao[1][2] = list_gangtao[1][2] + sheet.row_values(i)[g_col_kufangxiaoshou] + sheet.row_values(i)[g_col_kufangqita]
                        list_gangtao[1][3] += sheet.row_values(i)[g_col_kufangjiecun]

        print('档次', '记录数', '销量', '结存')

        for i in range(0, num_dangci):
            if (list_statics[i][0] == '40常规'):
                print(list_statics[i])
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                break
            elif (i == num_dangci - 1 and list_statics[i][0] != '40常规'):
                newtalbe.write(g_write_row_index, 0, '40常规')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)
        g_write_row_index += 1

        for i in range(0, num_dangci):
            if (list_statics[i][0] == '50常规'):
                print(list_statics[i])
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                break
            elif (i == num_dangci - 1 and list_statics[i][0] != '50常规'):
                newtalbe.write(g_write_row_index, 0, '50常规')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)
        g_write_row_index += 1

        for i in range(0, num_dangci):
            if (list_statics[i][0] == '60常规'):
                print(list_statics[i])
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                break
            elif (i == num_dangci - 1 and list_statics[i][0] != '60常规'):
                newtalbe.write(g_write_row_index, 0, '60常规')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)
        g_write_row_index += 1

        for i in range(0, num_dangci):
            if (list_statics[i][0] == '70常规'):
                print(list_statics[i])
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                break
            elif (i == num_dangci - 1 and list_statics[i][0] != '70常规'):
                newtalbe.write(g_write_row_index, 0, '70常规')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)
        g_write_row_index += 1

        for i in range(0, num_dangci):
            if (list_statics[i][0] == '80常规'):
                print(list_statics[i])
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                break
            elif (i == num_dangci - 1 and list_statics[i][0] != '80常规'):
                newtalbe.write(g_write_row_index, 0, '80常规')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)
        g_write_row_index += 1

        for i in range(0, num_dangci):
            if (list_statics[i][0] == '90常规'):
                print(list_statics[i])
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                break
            elif (i == num_dangci - 1 and list_statics[i][0] != '90常规'):
                newtalbe.write(g_write_row_index, 0, '90常规')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)
        g_write_row_index += 1

        for i in range(0, num_dangci):
            if (list_statics[i][0] == '非标准丁级'):
                print(list_statics[i])
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                break
            elif (i == num_dangci - 1 and list_statics[i][0] != '非标准丁级'):
                newtalbe.write(g_write_row_index, 0, '非标准丁级')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)
        g_write_row_index += 1

        for i in range(0, num_dangci):
            if (list_statics[i][0] == '非标准甲级'):
                print(list_statics[i])
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                break
            elif (i == num_dangci - 1 and list_statics[i][0] != '非标准甲级'):
                newtalbe.write(g_write_row_index, 0, '非标准甲级')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)
        g_write_row_index += 1

        for i in range(0, num_dangci):
            if (list_statics[i][0] == '丁级'):
                print(list_statics[i])
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                break
            elif (i == num_dangci - 1 and list_statics[i][0] != '丁级'):
                newtalbe.write(g_write_row_index, 0, '丁级')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)
        g_write_row_index += 1

        for i in range(0, num_dangci):
            if (list_statics[i][0] == '甲级'):
                print(list_statics[i])
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                break
            elif (i == num_dangci - 1 and list_statics[i][0] != '甲级'):
                newtalbe.write(g_write_row_index, 0, '甲级')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)
        g_write_row_index += 1

        for i in range(0, num_dangci):
            if (list_statics[i][0] == '钢套门'):
                print('70钢套门', list_gangtao[0][1], list_gangtao[0][2], list_gangtao[0][3])
                newtalbe.write(g_write_row_index, 0, '70钢套门')
                newtalbe.write(g_write_row_index, 1, list_gangtao[0][1])
                newtalbe.write(g_write_row_index, 2, list_gangtao[0][2])
                newtalbe.write(g_write_row_index, 3, list_gangtao[0][3])
                g_write_row_index += 1

                print('90钢套门', list_gangtao[1][1], list_gangtao[1][2], list_gangtao[1][3])
                newtalbe.write(g_write_row_index, 0, '90钢套门')
                newtalbe.write(g_write_row_index, 1, list_gangtao[1][1])
                newtalbe.write(g_write_row_index, 2, list_gangtao[1][2])
                newtalbe.write(g_write_row_index, 3, list_gangtao[1][3])
                g_write_row_index += 1

                print('钢套门合计', list_gangtao[0][1] + list_gangtao[1][1], list_gangtao[0][2] + list_gangtao[1][2],
                      list_gangtao[0][3] + list_gangtao[1][3])
                newtalbe.write(g_write_row_index, 0, '钢套门合计')
                newtalbe.write(g_write_row_index, 1, list_gangtao[0][1] + list_gangtao[1][1])
                newtalbe.write(g_write_row_index, 2, list_gangtao[0][2] + list_gangtao[1][2])
                newtalbe.write(g_write_row_index, 3, list_gangtao[0][3] + list_gangtao[1][3])
                g_write_row_index += 1
                break
            elif (i == num_dangci - 1 and list_statics[i][0] != '钢套门'):
                print('钢套门合计', 0, 0, 0)
                newtalbe.write(g_write_row_index, 0, '钢套门合计')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)

        print('外购', list_gangtao[2][1], list_gangtao[2][2], list_gangtao[2][3])

        for i in range(0, num_dangci):
            if (list_statics[i][0] == '外购'):
                print('外购', list_gangtao[2][1], list_gangtao[2][2], list_gangtao[2][3])
                newtalbe.write(g_write_row_index, 0, '外购门')
                newtalbe.write(g_write_row_index, 1, list_gangtao[2][1])
                newtalbe.write(g_write_row_index, 2, list_gangtao[2][2])
                newtalbe.write(g_write_row_index, 3, list_gangtao[2][3])
                g_write_row_index += 1

                break
            elif (i == num_dangci - 1 and list_statics[i][0] != '外购'):
                print('外购', 0, 0, 0)
                newtalbe.write(g_write_row_index, 0, '外购')
                newtalbe.write(g_write_row_index, 1, 0)
                newtalbe.write(g_write_row_index, 2, 0)
                newtalbe.write(g_write_row_index, 3, 0)
                g_write_row_index += 1

        print('**********档次统计-结束**********')

        print('\n**********规格统计-开始**********')
        g_write_row_index += 1
        # 获取开向种类数
        list_guige = [0] * (nrows - g_row_data)

        for i in range(g_row_data, nrows):
            if (sheet.row_values(i)[2].find('锁') == -1):  # find-在字符串中查找某子串，找不到则返回-1，此处只统计非锁的项目
                list_guige[i - g_row_data] = sheet.row_values(i)[g_col_guige]

        list_guige_count = set(list_guige)
        num_guige = len(list_guige_count)

        list_statics = [[0] * 4 for i in range(num_guige)]  # 4列分别为：规格、数据条数、销量、结存
        i = 0
        for name in list_guige_count:
            list_statics[i][0] = name
            i += 1

        # 处理原始数据，统计每种开向的销量和结存信息
        for i in range(g_row_data, nrows):
            for j in range(0, num_guige):
                if (list_statics[j][0] == sheet.row_values(i)[g_col_guige]):
                    list_statics[j][1] += 1
                    list_statics[j][2] = list_statics[j][2] + sheet.row_values(i)[g_col_kufangxiaoshou] + sheet.row_values(i)[g_col_kufangqita]
                    list_statics[j][3] += sheet.row_values(i)[g_col_kufangjiecun]

        list_statics.sort(key=lambda x: x[2], reverse=True)  # 按销量从高到低排序
        print('规格', '记录数', '销量', '结存')
        for i in range(0, num_guige):
            if (list_statics[i][0] != '' and list_statics[i][0] != 0):
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                g_write_row_index += 1
                print(list_statics[i])

        print('**********规格统计-结束**********')

        print('\n**********花色统计-开始**********')
        g_write_row_index += 1
        # 获取花色种类数
        list_huase = [0] * (nrows - g_row_data)

        for i in range(g_row_data, nrows):
            if (sheet.row_values(i)[2].find('锁') == -1):  # find-在字符串中查找某子串，找不到则返回-1，此处只统计非锁的项目
                list_huase[i - g_row_data] = sheet.row_values(i)[g_col_huase]

        list_huase_count = set(list_huase)
        num_huase = len(list_huase_count)

        list_statics = [[0] * 4 for i in range(num_huase)]  # 4列分别为：规格、数据条数、销量、结存
        i = 0
        for name in list_huase_count:
            list_statics[i][0] = name
            i += 1

        # 处理原始数据，统计每种开向的销量和结存信息
        for i in range(g_row_data, nrows):
            for j in range(0, num_huase):
                if (list_statics[j][0] == sheet.row_values(i)[g_col_huase]):
                    list_statics[j][1] += 1
                    list_statics[j][2] = list_statics[j][2] + sheet.row_values(i)[g_col_kufangxiaoshou] + sheet.row_values(i)[g_col_kufangqita]
                    list_statics[j][3] += sheet.row_values(i)[g_col_kufangjiecun]

        list_statics.sort(key=lambda x: x[2], reverse=True)  # 按销量从高到低排序

        print('开向', '记录数', '销量', '结存')
        for i in range(0, num_huase):
            if (list_statics[i][0] != '' and list_statics[i][0] != 0):
                for j in range(0, 4):
                    newtalbe.write(g_write_row_index, j, list_statics[i][j])
                g_write_row_index += 1
                print(list_statics[i])
        print('**********花色统计-结束**********')

        dstfile.save(destFullName)
