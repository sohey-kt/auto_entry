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
    # tk.messagebox.showinfo('ファイルを選択', 'OSAU フォーマットが記載されているExcelファイルを選択してください。')
    fileType = [('Excelファイル','*.xlsx'), ('Excelファイル', '*.xls')] 
    iniDir = os.path.abspath(os.path.dirname(__file__))
    filepath = tk.filedialog.askopenfilename(filetypes=fileType, initialdir = iniDir)

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

    #行列確認用
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
    # -1行目は生年月日など同じ番号の欄が複数あるもの。これは配列に格納する。
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
    return excel_info


def run_entry(excel_info):
    try:
        driver = webdriver.Chrome()
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
        apply_year = driver.find_element_by_id('UP1370_campaignApplyYear')
        apply_year.clear()
        apply_month = driver.find_element_by_id('UP1370_campaignApplyMonth')
        apply_month.clear()
        apply_date = driver.find_element_by_id('UP1370_campaignApplyDate')
        apply_date.clear()
        apply_year.send_keys(int(excel_info['2'][0]))
        apply_month.send_keys(int(excel_info['2'][1]))
        apply_date.send_keys(int(excel_info['2'][2]))
        driver.find_element_by_id('UISST0260_next').click()
        sleep(2)

        #ご利用エリア
            #固定電話番号
        fixed_tel = excel_info['3'].split('-')
        driver.find_element_by_id('UP1020_telNo1').send_keys(fixed_tel[0])
        driver.find_element_by_id('UP1020_telNo2').send_keys(fixed_tel[1])
        driver.find_element_by_id('UP1020_telNo3').send_keys(fixed_tel[2])
            #郵便番号
        zip_num = excel_info['4'].split('-')
        driver.find_element_by_id('UP1020_zipCd1').send_keys(zip_num[0])
        driver.find_element_by_id('UP1020_zipCd2').send_keys(zip_num[1])
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
            tk.messagebox.showwarning('注意', 'auプランの記載に誤りがある可能性があります')
        sleep(2)

        #住所選択
            #住所選択
        address_list = []
        soup = BeautifulSoup(driver.page_source, 'lxml')
        trs = soup.select('#UISST0090 > div.d-article > table > tbody > tr')
        zennkaku_hannkaku = {'０': '0', '１': '1', '２': '2', '３': '3', '４': '4', '５': '5', '６':'6', '７': '7', '８': '8', '９': '9'}
        for tr in trs:
            text = ''
            td_label = tr.select_one('td > label').text.replace('\n', '').replace('\t', '')
            for moji in td_label:
                zennkaku = re.match('[０-９]', moji)
                if zennkaku:
                    moji = zennkaku_hannkaku[moji]
                text += moji
            address_list.append(text)
        count = 0
        my_address = '〒' + str(excel_info['4']) + excel_info['32'][0] + excel_info['32'][1]
        for address in address_list:
            if address == my_address:
                break
            else:
                count += 1
        try:
            driver.find_element_by_id(f'UP1110_address{count}').click()
        except:
            tk.messagebox.showwarning('注意', '住所を見つけられませんでした、該当の住所があれば選択してお待ちください。\n10秒後に処理が再開します。')
            sleep(10)
        driver.find_element_by_id('UISST0090_next').click()
        sleep(2)

        #提供判定
        try:
            driver.find_element_by_id('UISST0150_next').click()
            sleep(2)
        except:
            tk.messagebox.showwarning('注意', 'ご指定の住所では、ご希望のコースをお申し込みいただけません')

        #オプション
        option_dict = {'10': '0', '8': '1', '16': '2', '41': '3', '12': '4', '13': '5', '11': '6', '9': '7', '15': '9', '3': '10', '17': '11'}
        for format_num in option_dict:
            if excel_info[format_num] == '○' or (format_num == '41' and excel_info['41'] == '5G') or (format_num == '3' and (excel_info['3'] or excel_info['6'] == '新番発番')):
                driver.find_element_by_id(f'UP1315_add_option_{option_dict[format_num]}_btn').click()
                sleep(3)
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
                    driver.find_element_by_id('UP1240_selectOption_0_0').click()
                    driver.find_element_by_id('UP1240_selectOption_1_0').click()
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
                        tk.messagebox.showwarning('注意', '現在電話サービスの欄を入力してください')
                    #下の入力欄
                    if excel_info['20'] == '掲載':
                        driver.find_element_by_id('UP1230_selectOption_0_1').click()
                    elif excel_info['20'] == '非掲載':
                        driver.find_element_by_id('UP1230_selectOption_0_2').click()
                    else:
                        tk.messagebox.showwarning('注意', '掲載/非掲載の欄を入力してください')
                    driver.find_element_by_id('UP1230_selectOption_1_0').click()
                    if excel_info['14'] == '解約' or excel_info['14'] == '継続':
                        driver.find_element_by_id('UP1230_selectOption_AU0103002').click()
                        driver.find_element_by_id('UP1230_selectOption_AU0103003').click()
                    elif excel_info['14'] == 'なし':
                        num = 21
                        service_ids = ['3', '2', '6', '4', '5', '9']
                        for service_id in service_ids:
                            if excel_info[str(num)] == '●':
                                driver.find_element_by_id(f'UP1230_selectOption_AU010300{service_id}')
                            else:
                                pass
                            num += 1
                    else:
                        tk.messagebox.showwarning('注意', '電話OPEX欄を入力してください')
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
                sleep(7)
        driver.find_element_by_id('submit').click()
        sleep(7)

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
            tk.messagebox.showwarning('注意', '性別の欄を入力してください')
        select_year = Select(driver.find_element_by_id('UP2010_birthYearKind_Year'))
        select_year.select_by_value(str(int(excel_info['31'][0])))
        select_month = Select(driver.find_element_by_id('UP2010_birthYearKind_Month'))
        select_month.select_by_value(str(int(excel_info['31'][1])))
        select_day = Select(driver.find_element_by_id('UP2010_birthYearKind_Day'))
        select_day.select_by_value(str(int(excel_info['31'][2])))
            #住所(入会証送付先)
        driver.find_element_by_id('UP2010_usrAddrBlock1').send_keys(int(excel_info['32'][2]))
        driver.find_element_by_id('UP2010_usrAddrBlock2').send_keys(int(excel_info['32'][3]))
        if excel_info['33']:
            driver.find_element_by_id('UP2010_edaban_input').click()
            driver.find_element_by_id('UP2010_usrAddrBlock3').send_keys(int(excel_info['33']))
        if excel_info['34']:
            driver.find_element_by_id('UP2010_usrAddrBuildingName').send_keys(excel_info['34'])
        if excel_info['35']:
            if type(excel_info['35']) == float:
                driver.find_element_by_id('UP2010_usrAddrRoomNo').send_keys(str(int(excel_info['35'])))
            elif type(excel_info['35']) == str:
                driver.find_element_by_id('UP2010_usrAddrRoomNo').send_keys(excel_info['35'])
            else:
                tk.messagebox.showwarning('注意', '部屋番号を数字か文字列で入力してください')
                print(type(excel_info['35']))
            #ご連絡先
        if excel_info['36']:
            driver.find_element_by_id('UP2010_contactTelKbn1').click()
            cellphone_tel = excel_info['36'].split('-')
            for i in range(3):
                driver.find_element_by_id(f'UP2010_contactTelNo{str(i + 1)}').send_keys(cellphone_tel[i])
            #お支払方法の登録・変更
        driver.find_element_by_id('UP2030_paymentKindKbn_kddiSeikyu').click()
        sleep(1.1)
        driver.find_element_by_id('UP2030_kddiPaymentKindKbn_notSum').click()
        sleep(1.1)
        driver.find_element_by_id('UP2030_kddiChrg_otherBillSendUmuFlg_1').click()
            #So-net 光 (auひかり) お申し込み者
        driver.find_element_by_id('UP2230_applyFamilyNameKnj').send_keys(excel_info['28'][0])
        driver.find_element_by_id('UP2230_applyFirstNameKnj').send_keys(excel_info['28'][1])
        driver.find_element_by_id('UP2230_applyFamilyNameKana').send_keys(excel_info['29'][0])
        driver.find_element_by_id('UP2230_applyFirstNameKana').send_keys(excel_info['29'][1])
            #ご利用場所住所
        driver.find_element_by_id('UP2230_block1').send_keys(int(excel_info['32'][2]))
        driver.find_element_by_id('UP2230_block2').send_keys(int(excel_info['32'][3]))
        if excel_info['33']:
            driver.find_element_by_id('UP2230_edaban_input').click()
            driver.find_element_by_id('UP2230_block3').send_keys(int(excel_info['33']))
        if excel_info['34']:
            driver.find_element_by_id('UP2230_buildingName').send_keys(excel_info['34'])
        if excel_info['35']:
            if type(excel_info['35']) == float:
                driver.find_element_by_id('UP2230_roomNo_2').send_keys(str(int(excel_info['35'])))
            elif type(excel_info['35']) == str:
                driver.find_element_by_id('UP2230_roomNo_2').send_keys(excel_info['35'])
            else:
                tk.messagebox.showwarning('注意', '部屋番号を数字か文字列で入力してください')
        if excel_info['37'] == '戸建て':
            driver.find_element_by_id('UP2230_dwellingForm_0').click()
        elif excel_info['37'] == '集合住宅':
            driver.find_element_by_id('UP2230_dwellingForm_1').click()
        else:
            tk.messagebox.showerror('エラー', '不明なエラーが発生しました')
        if excel_info['38'] == '持家':
            driver.find_element_by_id('UP2230_ownershipKbn_0').click()
        elif excel_info['38'] == '賃貸':
            driver.find_element_by_id('UP2230_ownershipKbn_1').click()
        else:
            tk.messagebox.showerror('エラー', '不明なエラーが発生しました')
            #auスマートバリューの申し込み
        soup = BeautifulSoup(driver.page_source, 'lxml')
        auSmartValue = soup.select_one('#UP2230_auSmartValueApply_0')
        if auSmartValue:
            if excel_info['42'] == 'SV':
                driver.find_element_by_id('UP2230_auSmartValueApply_0').click()
            elif excel_info['42'] == '':
                driver.find_element_by_id('UP2230_auSmartValueApply_1').click()
            else:
                tk.messagebox.showwarning('注意', 'SV有無は「SV」又は空欄で記載してください。')
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