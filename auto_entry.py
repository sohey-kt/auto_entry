import os
import signal
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import xlrd
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import lxml
import re
from time import sleep


def select_button(filepath):
    fileType = [('Excelファイル','*.xlsx'), ('Excelファイル', '*.xls')] 
    iniDir = os.path.abspath(os.path.dirname(__file__))
    filepath = tk.filedialog.askopenfilename(filetypes=fileType, initialdir = iniDir)

    if filepath != '':
        #選択したファイルを表示
        label2 = tk.Label(text=f'↓\n\n選択したファイル :\n{filepath}\n\n{"↓":52s}{"↓":>54s}')
        label2.pack()

        #実行ボタンを表示
        entryButton = tk.Button(root, text='実行', command=lambda: entry_button(filepath))
        entryButton.pack(ipadx=20, padx=(100, 0), side='left')

        #やり直しボタンを表示
        reSelectButton = tk.Button(root, text='やり直す', command=lambda: re_select_button(label2, entryButton, reSelectButton))
        reSelectButton.pack(ipadx=15, padx=(0, 100), side='right')


def re_select_button(label2, entryButton, reSelectButton):
    label2.pack_forget()
    entryButton.pack_forget()
    reSelectButton.pack_forget()


def entry_button(filepath):
    #選択したExcelのデータを読み込む → 辞書に格納
    excel_info = read_excelFile(filepath)

    #申し込みを開始
    run_entry(excel_info)


def read_excelFile(filepath):
    excel_info = {}
    wb = xlrd.open_workbook(filepath)
    sheet = wb.sheets()[0]

    # #行列確認用
    # foo = [sheet.row_values(x) for x in range(37)]
    # cn = -1
    # for u in foo:
    #     cn += 1
    #     list_ = []
    #     cnn = 0
    #     for i in u:
    #         i = str(cnn) + str(i)
    #         list_.append(i)
    #         cnn += 1
    #     print(str(cn) + '  ' + str(list_))

    #[[辞書番号, 何行目か, 何列目か], [..., ..., ...], ...]
    # -1行目となっているのは生年月日など同じ番号の欄が複数あるもの。これは配列に格納する。
    #42番号目としてSVの有無を取得する
    pointers = [
        [1, 29, 5],
        [2, -1, [[1, 9], [1, 10], [1, 11]]],
        [3, 12, 1],
        [4, 5, 0],
        [5, 33, 5],
        [6, 14, 1],
        [8, 21, 5],
        [9, 23, 5],
        [10, 25, 5],
        [11, 21, 7],
        [12, 23, 7],
        [13, 25, 7],
        [14, 12, 6],
        [15, 14, 6],
        [16, 16, 6],
        [17, 35, 7],
        [18, -1, [[2, 0], [2, 2]]],
        [19, -1, [[1, 0], [1, 2]]],
        [20, 16, 9],
        [21, 21, 1],
        [22, 23, 1],
        [23, 25, 1],
        [24, 21, 3],
        [25, 23, 3],
        [26, 25, 3],
        [27, 16, 1],
        [28, -1, [[2, 4], [2, 6]]],
        [29, -1, [[1, 4], [1, 6]]],
        [30, 1, 8],
        [31, -1, [[10, 5], [10, 6], [10, 7]]],
        [32, -1, [[5, 2], [5, 4], [8, 0], [8, 1]]],
        [33, 8, 2],
        [34, 8, 3],
        [35, 8, 5],
        [36, 18, 1],
        [37, 10, 0],
        [38, 10, 2],
        [39, 8, 6],
        [40, 18, 5],
        [41, 31, 5],
        [42, 35, 5]
    ]
    for pointer in pointers:
        if pointer[1] == -1:
            ary = []
            for subpointer in pointer[2]:
                ary.append(sheet.cell_value(subpointer[0], subpointer[1]))
            excel_info[str(pointer[0])] = ary
        else:
            excel_info[str(pointer[0])] = sheet.cell_value(pointer[1], pointer[2])
    #電話番号のハイフンがなければ追加
    tel_format = ['3','4', '27', '36']
    for num in tel_format:
        hyphen = excel_info[num].find('-')
        if hyphen == -1:
            with_hyphen = ''
            count = 0
            if len(excel_info[num]) == 10:
                for i in list(str(excel_info[num])):
                    if count == 3 or count == 6:
                        with_hyphen += '-'
                    with_hyphen += i
                    count += 1
            elif len(excel_info[num]) == 11:
                for i in list(str(excel_info[num])):
                    if count == 3 or count == 7:
                        with_hyphen += '-'
                    with_hyphen += i
                    count += 1
            elif len(excel_info[num]) == 7:
                for i in list(str(excel_info[num])):
                    if count == 3:
                        with_hyphen += '-'
                    with_hyphen += i
                    count += 1
            else:
                pass
            excel_info[num] = with_hyphen
        else:
            pass
    
    print(excel_info)
    return excel_info


def run_entry(excel_info):
    try:
        driver = webdriver.Chrome(executable_path=r'C:\Users\sohei\Desktop\dev\sohei_office\sohei\auto_entry\chromedriver.exe')
    except:
        try:
            driver = webdriver.Chrome(executable_path=r'\\192.168.5.151\共有フォルダ\CALL FORCE株式会社\本店カスタマーサクセス部\バックオフィス\ＯＳＡＵ\実行シート\≫自動エントリー\chromedriver_win32\chromedriver.exe')
        except:
            iniDir = str(os.path.abspath(os.path.dirname(__file__)))
            pathDir = iniDir[:iniDir.find('自動エントリー') + 8] + r'chromedriver_win32\chromedriver.exe'
            try:
                driver = webdriver.Chrome(executable_path=pathDir)
            except:
                pathDir = 'Z:' + pathDir
                try:
                    driver = webdriver.Chrome(executable_path=pathDir)
                except:
                    try:
                        driver = webdriver.Chrome(executable_path=r'\\192.168.5.151\共有フォルダ\CALL FORCE株式会社\本店カスタマーサクセス部\バックオフィス\ＯＳＡＵ\実行シート\自動エントリー\chromedriver_win32\chromedriver')
                    except:
                        try:
                            driver = webdriver.Chrome(executable_path=r'\\192.168.5.151\共有フォルダ\CALL FORCE株式会社\本店カスタマーサクセス部\バックオフィス\ＯＳＡＵ\実行シート\自動エントリー\chromedriver_win32\chromedriver.exe')
                        except:
                            tk.messagebox.showerror('エラー', 'ブラウザを開けませんでした。')
    try:
        driver.get('https://www.so-net.ne.jp/signup/sst/UISST0290.xhtml')
        sleep(2)
        
        #ログイン
            #ID
        driver.find_element_by_id('UP1390_loginId').send_keys('onestop1407')
            #パスワード
        driver.find_element_by_id('UP1390_password').send_keys('20140701os')
        driver.find_element_by_id('UISST0290_login').click()
        sleep(2)

        #代理店コード
            #代理店コード
        agent_code = driver.find_element_by_id('UP1360_agentCd')
        if excel_info['1'] == 'なし':
            agent_code.send_keys('6AX001')
        elif excel_info['1'] == '1万円CB':
            agent_code.send_keys('6AX301')
        elif excel_info['1'] == '3万円CB':
            agent_code.send_keys('6AX501')
        else:
            tk.messagebox.showwarning('注意', 'ISPCPの欄の記載に誤りがある可能性があります')
        driver.find_element_by_id('UP1360_confirm').click()
        sleep(1)
            #キャンペーン適用基準日
        year_month_date = ['Year', 'Month', 'Date']
        for i, when in enumerate(year_month_date):
            apply_when = driver.find_element_by_id(f'UP1370_campaignApply{when}')
            apply_when.clear()
            apply_when.send_keys(int(excel_info['2'][i]))

        driver.find_element_by_id('UISST0260_next').click()
        sleep(2)

        #ご利用エリア
            #固定電話番号
        fixed_tel = excel_info['3'].split('-')
        try:
            for i in range(3):
                driver.find_element_by_id(f'UP1020_telNo{str(i + 1)}').send_keys(fixed_tel[i])
        except:
            tk.messagebox.showwarning('注意', '電話番号を正しく入力し直してください。\n10秒後に処理が再開します。')
            sleep(10)
            #郵便番号
        zip_num = excel_info['4'].split('-')
        try:
            for i in range(2):
                driver.find_element_by_id(f'UP1020_zipCd{str(i + 1)}').send_keys(zip_num[i])
        except:
            tk.messagebox.showwarning('注意', '郵便番号を正しく入力しなおしてください。\n10秒後に処理が再開します。')
            sleep(10)
        driver.find_element_by_id('UISST0040_next').click()
        sleep(2)

        #コース選択
        if excel_info['5'] == 'ずっとギガ得':
            driver.find_element_by_id('UP1060Useable_entry_AUHK0103').click()
        elif excel_info['5'] == 'マンションお得A':
            driver.find_element_by_id('UP1060Useable_entry_AUHK0202').click()
        elif excel_info['5'] == 'マンション標準':
            driver.find_element_by_id('UP1060Useable_entry_AUHK0201').click()
        else:
            tk.messagebox.showwarning('注意', 'auプランの欄を入力してください。\n10秒後に処理が再開します。')
            sleep(10)
        sleep(2)

        #住所選択
            #住所選択
        address_list = []
        soup = BeautifulSoup(driver.page_source, 'lxml')
                #ボタンのid名を取得
        if excel_info['5'] == 'ずっとギガ得':
            number_of_UISST = '09'
            name_of_UP = 'UP1110_address'
        elif excel_info['5'] == 'マンションお得A' or excel_info['5'] == 'マンション標準':
            number_of_UISST = '11'
            name_of_UP = 'UP1130_building_'
        else:
            tk.messagebox.showerror('エラー', 'auプランの欄の記載が不正です。')
                #住所選択肢の文字の補正
        trs = soup.select(f'#UISST0{number_of_UISST}0 > div.d-article > table > tbody > tr')
        zennkaku_hannkaku = {'０': '0', '１': '1', '２': '2', '３': '3', '４': '4', '５': '5', '６':'6', '７': '7', '８': '8', '９': '9'}
        for tr in trs:
            text = ''
            td_label = tr.select_one('td > label').text.replace('\n', '').replace('\t', '')
            for moji in td_label:
                zennkaku = re.match('[０-９]', moji)
                if zennkaku:
                    moji = zennkaku_hannkaku[moji]
                elif moji == 'ー' or moji == '－':
                    moji = '-'
                text += moji
            address_list.append(text)
        my_address = '〒' + str(excel_info['4'])
        for i in range(4):
            if i >= 2 and excel_info['32'][i]:
                my_address += '-'
                my_address += str(int(excel_info['32'][i]))
            else:
                my_address += excel_info['32'][i]
                #枝番→号→...の順で値を外していき、一致するまで検索していく
        match_address = False
        for _ in range(4):
            for count, address in enumerate(address_list):
                if address == my_address:
                    match_address = True
                    break
                if count == len(address_list) - 1:
                    hyphen = my_address.rfind('-')
                    if hyphen < 8:
                        #「丁目」があれば外す
                        if my_address[-2:] == '丁目':
                            my_address = my_address[:-2]
                        #丁目の数字を外す
                        for string in reversed(my_address):
                            integer = re.match(r'\d', string)
                            if integer:
                                my_address = my_address[:-1]
                            else:
                                break
                    else:
                        my_address = my_address[:hyphen]
            else:
                continue
            break

        if match_address:
            driver.find_element_by_id(name_of_UP + str(count)).click()
        else:
            tk.messagebox.showwarning('注意', '正しい住所を見つけられませんでした、\n該当の住所があれば選択して、次のページへ進むボタンを押す直前でお待ちください。\n13秒後に処理が再開します。')
            sleep(13)
        driver.find_element_by_id(f'UISST0{number_of_UISST}0_next').click()
        sleep(2)

        #自前エリアでの番地と号の入力(正確には恐らく自前エリアでSV付帯→5G,自前エリアでSV付帯無し→1G)
        soup = BeautifulSoup(driver.page_source, 'lxml')
        UP1150_block1 = soup.select_one('#UP1150_block1')
        if UP1150_block1:
            if type(excel_info['32'][2]) == float:
                driver.find_element_by_id('UP1150_block1').send_keys(int(excel_info['32'][2]))
            else:
                driver.find_element_by_id('UP1150_block1').send_keys(excel_info['32'][2])
            if excel_info['32'][3]:
                if type(excel_info['32'][3]) == float:
                    driver.find_element_by_id('UP1150_block2').send_keys(int(excel_info['32'][3]))
                else:
                    driver.find_element_by_id('UP1150_block2').send_keys(excel_info['32'][3])
            driver.find_element_by_id('UISST0130_next').click()
            sleep(2)

        #提供判定
        try:
            driver.find_element_by_id('UISST0150_next').click()
            sleep(3.5)
        except:
            tk.messagebox.showwarning('注意', 'ご指定の住所では、ご希望のコースをお申し込みいただけません')

        #オプション
        option_dict_url = {
            '3': 'https://www.so-net.ne.jp/access/hikari/au/phone/',
            '8': 'https://www.so-net.ne.jp/lifesupport/anshin/',
            '9': 'https://www.so-net.ne.jp/option/security/kaspersky/',
            '10': 'https://www.so-net.ne.jp/option/kurashi/omamori-wide/',
            '11': 'https://www.so-net.ne.jp/option/security/sagiwall/',
            '12': 'https://www.so-net.ne.jp/option/benefit/elavel-club/',
            '13': 'https://www.so-net.ne.jp/option/aosboxcool/index.html',
            '15': 'https://www.so-net.ne.jp/access/hikari/au/tv/',
            '16': 'https://www.so-net.ne.jp/option/visual/unext/',
            '17': 'https://www.so-net.ne.jp/access/hikari/au/musenlan/',
            '41': 'https://www.so-net.ne.jp/guide/au/highspeed.html'
        }
        option_dict_btn = {}
        soup = BeautifulSoup(driver.page_source, 'lxml')
        other_menu = soup.find('div', class_='other_menu')
        trs = other_menu.find_all('tr')
        for tr in trs:
            url = tr.a.get('href')
            for key, value in list(option_dict_url.items()):
                if value == url:
                    id_btn = tr.select_one('td:nth-child(2) > a').get('id')
                    option_dict_btn[key] = id_btn
                    option_dict_url.pop(key)

        for format_num in option_dict_btn:
            if excel_info[format_num] == '○' or (format_num == '41' and excel_info['41'] == '5G') or (format_num == '3' and (excel_info['3'] or excel_info['6'] == '新番発番')):
                sleep(2)
                driver.find_element_by_id(option_dict_btn[format_num]).click()
                sleep(1.5)
                #U-NEXT for So-net
                if format_num == '16':
                    driver.find_element_by_id('UP9220_guideUseFlg_1').click()
                #auひかり ホーム 高速サービス(5ギガ/10ギガ)
                elif format_num == '41':
                    driver.find_element_by_id('UP3007_selectOptionName_radio_0_0').click()
                #カスペルスキーセキュリティ
                elif format_num == '9':
                    driver.find_element_by_id('UIOPT0020_next').click()
                #auひかり テレビサービス
                elif format_num == '15':
                    STA3000 = driver.find_element_by_id('UP1240_selectOption_AU0201002')
                    if STA3000:
                        STA3000.click()
                    #オールジャンルパックを選択
                    soup = BeautifulSoup(driver.page_source, 'lxml')
                    trs = soup.select('#UISST0200 > div.d-article > table > tbody > tr')
                    for tr in trs:
                        td_label = tr.select_one('td > label').text.replace('\n', '').replace('\t', '')
                        if td_label == 'オールジャンルパック':
                            input_id = tr.select_one('td > input').get('id')
                            driver.find_element_by_id(input_id).click()
                            break
                    driver.find_element_by_id('UISST0200_entry').click()
                #auひかり 無線LANレンタル
                elif format_num == '17':
                    driver.find_element_by_id('UP1250_selectOption_0_0').click()
                    driver.find_element_by_id('UP1250_selectOption_AU0304001').click()
                    driver.find_element_by_id('UISST0210_entry').click()
                #auひかり 電話サービス
                elif format_num == '3':
                    #上の入力欄
                    if excel_info['6'] == '新番発番':
                        driver.find_element_by_id('UP1230_telNoGetKindKbn_1').click()
                    elif excel_info['3']:
                        driver.find_element_by_id('UP1230_telNoGetKindKbn_2').click()
                        sleep(1.1)
                        select_srvc = Select(driver.find_element_by_id('UP1230_usePhoneSrvc'))
                        if excel_info['6'] == 'NTTひかり電話':
                            select_srvc.select_by_value('0001')
                        elif excel_info['6'] == 'NUROひかり電話':
                            select_srvc.select_by_value('0901')
                        else:
                            tk.messagebox.showerror('エラー', '不明なエラーが発生しました')
                        for i in range(3):
                            driver.find_element_by_id(f'UP1230_keepUseTelNo{str(i + 1)}').send_keys(fixed_tel[i])
                        driver.find_element_by_id('UP1230_lineHolderFamilyNameKnj').send_keys(excel_info['18'][0])
                        driver.find_element_by_id('UP1230_lineHolderFirstNameKnj').send_keys(excel_info['18'][1])
                        driver.find_element_by_id('UP1230_lineHolderFamilyNameKana').send_keys(excel_info['19'][0])
                        driver.find_element_by_id('UP1230_lineHolderFirstNameKana').send_keys(excel_info['19'][1])
                    else:
                        tk.messagebox.showwarning('注意', '現在電話サービスの欄を入力してください。\n10秒後に処理が再開します。')
                        sleep(10)
                    #下の入力欄
                    if excel_info['20'] == '掲載':
                        driver.find_element_by_id('UP1230_selectOption_0_1').click()
                    elif excel_info['20'] == '非掲載' or excel_info['20'] == '':
                        driver.find_element_by_id('UP1230_selectOption_0_2').click()
                    else:
                        tk.messagebox.showwarning('注意', '掲載/非掲載の欄に不正な入力があります。\n10秒後に処理が再開します。')
                        sleep(10)
                    driver.find_element_by_id('UP1230_selectOption_1_0').click()
                    if excel_info['14'] == '解約' or excel_info['14'] == '継続':
                        driver.find_element_by_id('UP1230_selectOption_AU0103002').click()
                        driver.find_element_by_id('UP1230_selectOption_AU0103003').click()
                    elif excel_info['14'] == 'なし':
                        num = 21
                        service_ids = ['3', '2', '6', '4', '5', '9']
                        for service_id in service_ids:
                            if excel_info[str(num)] == '●':
                                if service_id == '9':
                                    try:
                                        driver.find_element_by_id('UP1230_selectOption_AU0103009').click()
                                    except:
                                        pass
                                else:
                                    driver.find_element_by_id(f'UP1230_selectOption_AU010300{service_id}').click()
                            else:
                                pass
                            num += 1
                    else:
                        tk.messagebox.showwarning('注意', '電話OPEX欄を入力してください。\n10秒後に処理が再開します。')
                        sleep(10)
                    #2番号目の申し込み
                    if excel_info['27']:
                        driver.find_element_by_id('Bangou21').click()
                        if excel_info['27'] == '新番発番':
                            driver.find_element_by_id('UP1230_telNoGetKindKbn_2_1').click()
                        else:
                            driver.find_element_by_id('UP1230_telNoGetKindKbn_2_2').click()
                            sleep(1.1)
                            select_srvc2 = Select(driver.find_element_by_id('UP1230_usePhoneSrvc2'))
                            if excel_info['6'] == 'NTTひかり電話':
                                select_srvc2.select_by_value('0001')
                            elif excel_info['6'] == 'NUROひかり電話':
                                select_srvc2.select_by_value('0901')
                            else:
                                tk.messagebox.showerror('エラー', '不明なエラーが発生しました')
                            fixed_tel2 = excel_info['27'].split('-')
                            for i in range(3):
                                driver.find_element_by_id(f'UP1230_keepUseTelNo2_{str(i + 1)}').send_keys(fixed_tel2[i])
                            driver.find_element_by_id('UP1230_lineHolderFamilyNameKnj2').send_keys(excel_info['18'][0])
                            driver.find_element_by_id('UP1230_lineHolderFirstNameKnj2').send_keys(excel_info['18'][1])
                            driver.find_element_by_id('UP1230_lineHolderFamilyNameKana2').send_keys(excel_info['19'][0])
                            driver.find_element_by_id('UP1230_lineHolderFirstNameKana2').send_keys(excel_info['19'][1])
                        driver.find_element_by_id('UP1230_selectOption_3_2').click()
                        driver.find_element_by_id('UP1230_selectOption_4_1').click()
                    else:
                        pass
                    driver.find_element_by_id('UISST0190_entry').click()
                #それ以外のサービス
                else:
                    pass
                #「次のページへ」ボタン
                if format_num != '15' and format_num != '17' and format_num != '9' and format_num != '3':
                    driver.find_element_by_id('UIOPT0030_next').click()
                sleep(2)
                    
        driver.find_element_by_id('submit').click()
        sleep(1.7)

        #入会情報
            #お名前
        driver.find_element_by_id('UP2010_usrFamilyNameKnj').send_keys(excel_info['28'][0])
        driver.find_element_by_id('UP2010_usrFirstNameKnj').send_keys(excel_info['28'][1])
        driver.find_element_by_id('UP2010_usrFamilyNameKana').send_keys(excel_info['29'][0])
        driver.find_element_by_id('UP2010_usrFirstNameKana').send_keys(excel_info['29'][1])
            #性別・生年月日
        if excel_info['30'] == '男':
            driver.find_element_by_id('UP2010_sex_0').click()
        elif excel_info['30'] == '女':
            driver.find_element_by_id('UP2010_sex_1').click()
        else:
            tk.messagebox.showwarning('注意', '性別の欄を入力してください\n10秒後に処理が再開します。')
            sleep(10)

        year_month_day = ['Year', 'Month', 'Day']
        for i, when in enumerate(year_month_day):
            select_when = Select(driver.find_element_by_id(f'UP2010_birthYearKind_{when}'))
            select_when.select_by_value(str(int(excel_info['31'][i])))

            #住所(入会証送付先)
                #自動入力されている部分が住所選択のワードのみと仮定して町名および丁目をただしく入力する
        UP2010_usrAddrCityName_value = driver.find_element_by_id('UP2010_usrAddrCityName').get_attribute('value')
        find_UP2010_usrAddrCityName_value = excel_info['32'][1].find(UP2010_usrAddrCityName_value)
        if find_UP2010_usrAddrCityName_value != -1:
            city_name_info = excel_info['32'][1][len(UP2010_usrAddrCityName_value):]
            UP2010_usrAddrTownName = driver.find_element_by_id('UP2010_usrAddrTownName')
            UP2010_usrAddrTownName.clear()
            UP2010_usrAddrTownName.send_keys(city_name_info)
        else:
            tk.messagebox.showerror('エラー', '獲得用紙と表示されている市区町村が一致しないため、入力を中止しました。')
                #番地入力確認
        soup = BeautifulSoup(driver.page_source, 'lxml')
        UP2230_buildingName = soup.find('input', id='UP2010_usrAddrBlock1')
        if UP2230_buildingName:
            try:
                UP2230_buildingName = UP2230_buildingName['value']
            except KeyError:
                if excel_info['32'][2]:
                    driver.find_element_by_id('UP2010_usrAddrBlock1').send_keys(int(excel_info['32'][2]))
                if excel_info['32'][3]:
                    driver.find_element_by_id('UP2010_usrAddrBlock2').send_keys(int(excel_info['32'][3]))
                if excel_info['33']:
                    driver.find_element_by_id('UP2010_edaban_input').click()
                    sleep(0.1)
                    driver.find_element_by_id('UP2010_usrAddrBlock3').send_keys(int(excel_info['33']))
        if excel_info['34']:
            UP2010_usrAddrBuildingName = soup.find('input', id='UP2010_usrAddrBuildingName')
            if UP2010_usrAddrBuildingName:
                try:
                    UP2010_usrAddrBuildingName = UP2010_usrAddrBuildingName['value']
                except KeyError:
                    driver.find_element_by_id('UP2010_usrAddrBuildingName').send_keys(excel_info['34'])
        if excel_info['35']:
            if type(excel_info['35']) == float:
                driver.find_element_by_id('UP2010_usrAddrRoomNo').send_keys(str(int(excel_info['35'])))
            elif type(excel_info['35']) == str or type(excel_info['35']) == int:
                driver.find_element_by_id('UP2010_usrAddrRoomNo').send_keys(excel_info['35'])
            else:
                tk.messagebox.showwarning('注意', '部屋番号が数字か文字列で記載されていませんでした。\n部屋番号を入力してください。\n10秒後に処理が再開します。')
                sleep(10)
            #ご連絡先
        if excel_info['36']:
            driver.find_element_by_id('UP2010_contactTelKbn1').click()
            sleep(0.2)
            cellphone_tel = excel_info['36'].split('-')
            for i in range(3):
                try:
                    driver.find_element_by_id(f'UP2010_contactTelNo{str(i + 1)}').send_keys(cellphone_tel[i])
                except:
                    tk.messagebox.showwarning('注意', '携帯電話番号を正しく入力し直してください。\n10秒後に処理が再開します。')
                    sleep(10)
        else:
            driver.find_element_by_id('UP2010_contactTelKbn0').click()
            #お支払方法の登録・変更
        driver.find_element_by_id('UP2030_paymentKindKbn_kddiSeikyu').click()
        sleep(1.3)
        driver.find_element_by_id('UP2030_kddiPaymentKindKbn_notSum').click()
        sleep(1.3)
        driver.find_element_by_id('UP2030_kddiChrg_otherBillSendUmuFlg_1').click()
            #So-net 光 (auひかり) お申し込み者
        driver.find_element_by_id('UP2230_applyFamilyNameKnj').send_keys(excel_info['28'][0])
        driver.find_element_by_id('UP2230_applyFirstNameKnj').send_keys(excel_info['28'][1])
        driver.find_element_by_id('UP2230_applyFamilyNameKana').send_keys(excel_info['29'][0])
        driver.find_element_by_id('UP2230_applyFirstNameKana').send_keys(excel_info['29'][1])
            #ご利用場所住所
                #記入枠があれば内容を削除して獲得用紙内の内容を記載
        UP2230_townName = soup.find('input', id='UP2230_townName')
        if UP2230_townName:
            UP2230_cityName_value = driver.find_element_by_id('UP2230_cityName').get_attribute('value')
            if UP2230_cityName_value:
                find_UP2230_cityName_value = excel_info['32'][1].find(UP2230_cityName_value)
                if find_UP2230_cityName_value != -1:
                    city_name_info2 = excel_info['32'][1][len(UP2230_cityName_value):]
                    UP2230_townName = driver.find_element_by_id('UP2230_townName')
                    UP2230_townName.clear()
                    UP2230_townName.send_keys(city_name_info2)
                else:
                    tk.messagebox.showerror('エラー', '獲得用紙と表示されている市区町村が一致しないため、入力を中止しました。')
            else:
                tk.messagebox.showerror('エラー', '市区町村が入力されていません。')
                #記入枠があるかつ空欄の場合に記入する
        if excel_info['32'][2]:
            UP2230_block1 = soup.find('input', id='UP2230_block1')
            if UP2230_block1:
                try:
                    UP2230_block1_value = UP2230_block1['value']
                except KeyError:
                    driver.find_element_by_id('UP2230_block1').send_keys(int(excel_info['32'][2]))
                    #番地が記入でき、かつ空欄なら号も記入できるのでここに置く
                    if excel_info['32'][3]:
                        driver.find_element_by_id('UP2230_block2').send_keys(int(excel_info['32'][3]))
                    #枝番も同様
                    if excel_info['33']:
                        driver.find_element_by_id('UP2230_edaban_input').click()
                        driver.find_element_by_id('UP2230_block3').send_keys(int(excel_info['33']))
                #記入枠があるかつ空欄の場合に記入する
        if excel_info['34']:
            UP2230_buildingName = soup.find('input', id='UP2230_buildingName')
            if UP2230_buildingName:
                try:
                    UP2230_buildingName_value = UP2230_buildingName['value']
                except KeyError:
                    driver.find_element_by_id('UP2230_buildingName').send_keys(excel_info['34'])

        if excel_info['35']:
            if type(excel_info['35']) == float:
                try:
                    driver.find_element_by_id('UP2230_roomNo_2').send_keys(str(int(excel_info['35'])))
                except:
                    try:
                        driver.find_element_by_id('UP2230_roomNo').send_keys(str(int(excel_info['35'])))
                    except:
                        tk.messagebox.showerror('エラー', '部屋番号入力でエラーが発生しました。')
            elif type(excel_info['35']) == str:
                try:
                    driver.find_element_by_id('UP2230_roomNo_2').send_keys(excel_info['35'])
                except:
                    try:
                        driver.find_element_by_id('UP2230_roomNo').send_keys(excel_info['35'])
                    except:
                        tk.messagebox.showerror('エラー', '部屋番号入力でエラーが発生しました。')
            else:
                tk.messagebox.showwarning('注意', '部屋番号が数字か文字列で記載されていませんでした。\n部屋番号を入力してください。\n10秒後に処理が再開します。')
                sleep(10)
        UP2230_dwellingForm = soup.select_one('#UP2230_dwellingForm_0')
        if UP2230_dwellingForm:
            if excel_info['37'] == '戸建て':
                driver.find_element_by_id('UP2230_dwellingForm_0').click()
            elif excel_info['37'] == '集合住宅':
                driver.find_element_by_id('UP2230_dwellingForm_1').click()
            else:
                tk.messagebox.showwarning('注意', '集合住宅か戸建てを選択してください。\n10秒後に処理が再開します。')
                sleep(10)
        floor = soup.select_one('#UP2230_floor')
        if floor:
            select_floor = Select(driver.find_element_by_id('UP2230_floor'))
            if excel_info['39']:
                floor_num = excel_info['39'][0]
                if type(floor_num) == float:
                    floor_num = int(floor_num)
                if str(floor_num) == '1' or str(floor_num) == '１':
                    select_floor.select_by_value('1')
                elif str(floor_num) == '2' or str(floor_num) == '２':
                    select_floor.select_by_value('2')
                else:
                    tk.messagebox.showwarning('注意', '1文字目に階数の数字記入されているか、\n利用階の階数が条件を満たしているかを確認してください。\n10秒後に処理が再開します。')
                    sleep(10)
        UP2230_ownershipKbn = soup.select_one('#UP2230_ownershipKbn_0')
        if UP2230_ownershipKbn:
            if excel_info['38'] == '持家':
                driver.find_element_by_id('UP2230_ownershipKbn_0').click()
            elif excel_info['38'] == '賃貸':
                driver.find_element_by_id('UP2230_ownershipKbn_1').click()
            else:
                tk.messagebox.showwarning('注意', '持家か賃貸を選択してください。')
            #現在ご利用中のインターネット回線
        soup = BeautifulSoup(driver.page_source, 'lxml')
        existLine = soup.select_one('#UP2230_existLineFlg_0')
        if existLine:
            if excel_info['40'] == 'NURO' or excel_info['40'] == '未設' or excel_info['40'] == 'FVNO':
                driver.find_element_by_id('UP2230_existLineFlg_0').click()
            elif excel_info['40'] == 'NTT':
                driver.find_element_by_id('UP2230_existLineFlg_1').click()
            elif excel_info['40'] == 'コラボ':
                driver.find_element_by_id('UP2230_existLineFlg_2').click()
            else:
                tk.messagebox.showwarning('注意', '現回線の欄を選択してください。')
            #他光回線からの切り替え
        UP2230_nowLineChangeKbn = soup.select_one('#UP2230_nowLineChangeKbn_0')
        if UP2230_nowLineChangeKbn:
            if excel_info['40'] == 'NTT':
                driver.find_element_by_id('UP2230_nowLineChangeKbn_1').click()
            elif excel_info['40'] == 'コラボ' or excel_info['40'] == 'NURO' or excel_info['40'] == 'FVNO':
                driver.find_element_by_id('UP2230_nowLineChangeKbn_2').click()
            elif excel_info['40'] == '未設':
                driver.find_element_by_id('UP2230_nowLineChangeKbn_0').click()
            else:
                tk.messagebox.showerror('エラー', '現回線の欄を記入してください。')

            #auスマートバリューの申し込み
        auSmartValue = soup.select_one('#UP2230_auSmartValueApply_0')
        if auSmartValue:
            if excel_info['42'] == 'SV':
                driver.find_element_by_id('UP2230_auSmartValueApply_0').click()
            elif excel_info['42'] == '':
                driver.find_element_by_id('UP2230_auSmartValueApply_1').click()
            else:
                tk.messagebox.showwarning('注意', 'SV有無は「SV」又は空欄で記載してください。/n10秒後に処理が再開します。')
                sleep(10)
        else:
            pass
            #KDDIからのお知らせ送付先
        driver.find_element_by_id('UP2240_sendAddressSelectFlg_0').click()
        driver.find_element_by_id('submit').click()
        sleep(5)
    finally:
        os.kill(driver.service.process.pid,signal.SIGTERM)


if __name__ == '__main__':
    filepath = ''
    root = tk.Tk()
    root.title('自動エントリーシステム')
    root.geometry('600x210')

    #選択ボタン
    selectButton = tk.Button(root, text='選択', command=lambda: select_button(filepath))
    selectButton.pack(ipadx=30, pady=10)

    root.mainloop()