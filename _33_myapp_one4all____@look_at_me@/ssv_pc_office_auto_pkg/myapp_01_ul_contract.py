'''
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
# generate contract automatically
# v0, @Z.Liang, 20190507
# v1, @Z.Liang, 20190729. added GUI
# v2, fix a potential server drive path problem
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
'''
import tkinter, os
from tkinter import (
    Toplevel,
    PhotoImage,
    Frame,
    Label,
    Entry,
    INSERT,
    Button,
    Text,
    END,
    messagebox,
)
from datetime import datetime

class UL_Gmbh_contract():
    BASE_DIR = os.path.dirname(os.path.abspath(os.path.dirname(__file__)))
    contract_year = datetime.now().year
    company_name = 'UL GmbH'
    follow_cost = 7020
    doc_cost = 3569
    other_cost = 0
    USD2RMB_forex = 6.47

    app_width = 800
    app_height = 620
    favicon_path = os.path.join(BASE_DIR, r"data\assets\favicon_pear.png")

    def __init__(self):
        ##<~ app_wm
        self.root = Toplevel()
        self.root.title("UL Gmbh contract generator v.0, by Z.Liang, 2019")
        self.root.geometry("{}x{}".format(self.app_width, self.app_height))
        try:
            self.imgicon = PhotoImage(file=self.favicon_path)
            self.root.tk.call('wm', 'iconphoto', self.root._w, self.imgicon)
        except tkinter._tkinter.TclError:
            pass
        self.root.focus_set()

        ##<~frame1
        self.fm1 = Frame(self.root)
        self.fm1.grid(row=0, column=0, padx=2, sticky="NW")

        self.lbl_contract_year = Label(self.fm1, text='contract_year:')
        self.lbl_contract_year.grid(row=0, column=0, sticky="W")
        self.lbl_company_name = Label(self.fm1, text='company_name:')
        self.lbl_company_name.grid(row=1, column=0, sticky="W")
        self.lbl_follow_cost = Label(self.fm1, text='follow_cost(USD):')
        self.lbl_follow_cost.grid(row=2, column=0, sticky="W")
        self.lbl_doc_cost = Label(self.fm1, text='doc_cost(USD):')
        self.lbl_doc_cost.grid(row=3, column=0, sticky="W")
        self.lbl_other_cost = Label(self.fm1, text='other_cost(USD):')
        self.lbl_other_cost.grid(row=4, column=0, sticky="W")
        self.lbl_USD2RMB_forex = Label(self.fm1, text='USD2RMB_forex:')
        self.lbl_USD2RMB_forex.grid(row=5, column=0, sticky="W")
        self.lbl_ttl_amt = Label(self.fm1, text='ttl_amt_USD:')
        self.lbl_ttl_amt.grid(row=6, column=0, sticky="W")
        self.lbl_ttl_amt_RMB = Label(self.fm1, text='ttl_amt_RMB:')
        self.lbl_ttl_amt_RMB.grid(row=7, column=0, sticky="W")

        self.e_contract_year = Entry(self.fm1, width=15)
        self.e_contract_year.insert(INSERT, self.contract_year)
        self.e_contract_year.grid(row=0, column=1, sticky="W")

        self.e_company_name = Entry(self.fm1, width=15)
        self.e_company_name.insert(INSERT, self.company_name)
        self.e_company_name.grid(row=1, column=1, sticky="W")

        self.e_follow_cost = Entry(self.fm1, width=15)
        self.e_follow_cost.insert(INSERT, 0)
        self.e_follow_cost.grid(row=2, column=1, sticky="W")

        self.e_doc_cost = Entry(self.fm1, width=15)
        self.e_doc_cost.insert(INSERT, 0)
        self.e_doc_cost.grid(row=3, column=1, sticky="W")

        self.e_other_cost = Entry(self.fm1, width=15)
        self.e_other_cost.insert(INSERT, 0)
        self.e_other_cost.grid(row=4, column=1, sticky="W")

        self.e_USD2RMB_forex = Entry(self.fm1, width=15)
        self.e_USD2RMB_forex.insert(INSERT, self.USD2RMB_forex)
        self.e_USD2RMB_forex.grid(row=5, column=1, sticky="W")

        self.lbl_ttl_amt = Label(self.fm1, width=15)
        self.lbl_ttl_amt['text'] = ''
        self.lbl_ttl_amt.grid(row=6, column=1, sticky="W")

        self.lbl_ttl_amt_RMB = Label(self.fm1, width=15)
        self.lbl_ttl_amt_RMB['text'] = ''
        self.lbl_ttl_amt_RMB.grid(row=7, column=1, sticky="W")

        ##<~ start button
        self.btn_start = Button(self.fm1, text='START', command=self.generate_flowlites_contract_content)
        self.btn_start.grid(row=8, column=0, padx=10, pady=10, sticky="NW")

        ##<~frame2
        self.fm2 = Frame(self.root)
        self.fm2.grid(row=0, column=1, padx=3, sticky="NW")
        self.txt_contract = Text(self.fm2, width=140, height=48)
        self.txt_contract.grid(row=0, column=0, sticky="W")
        self.txt_contract.insert(INSERT, "")

        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

    def generate_flowlites_contract_content(self):
        ## dynamically change total amount based on input
        self.e_USD2RMB_forex.update()
        new_ttl_amt = float(self.e_follow_cost.get()) + float(self.e_doc_cost.get()) + float(self.e_other_cost.get())
        new_ttl_amt_RMB = new_ttl_amt * float(self.e_USD2RMB_forex.get())

        self.lbl_ttl_amt['text'] = new_ttl_amt
        self.lbl_ttl_amt.update()
        self.lbl_ttl_amt_RMB['text'] = format(new_ttl_amt_RMB, "2f")
        self.lbl_ttl_amt.update()

        contract_year = self.e_contract_year.get().strip()
        company_name = self.e_company_name.get().strip()
        follow_cost = float(self.e_follow_cost.get().strip())
        doc_cost = float(self.e_doc_cost.get().strip())
        other_cost = float(self.e_other_cost.get().strip())
        ttl_amt = float(self.lbl_ttl_amt['text'])
        ttl_amt_RMB = float(self.lbl_ttl_amt_RMB['text'])

        flowlites_content = ('''
{0}年度{1}跟踪服务合同签订

目的: 用于支付{0}年{1}跟踪服务费
背景: 从2017/7/1起,中国税务局变更了外汇支付方针.根据新方针外汇支付前必须签订合同.
现状: {0}年度SSV与{1}还未签订合同.

SSV与UL检测瑞士公司({1})签订跟踪服务合同
合同期间：{0}年1月1日～{0}年12月31日
决裁金额：{4:,.2f}美金
费用明细:
1.跟踪服务费:{2:,.2f}美金
2.{0}年度档案维护费:{3:,.2f}美金
3.其他与跟踪服务相关的费用:{5:,.2f}美金
由于每年的跟踪服务费会变化,因此每年需要签订一次合同.

-------br-------br-------br-------br-------br-------

{0}年度{1}サービス費用契約書締結

目的: {0}年度{1}へのサービス費用支払い
背景: 2017/7/1より税務局から外貨の支払いについて、すべて契約締結が必要となった。
現状: {0}年度、{1}とのサービス費用契約書がまだ契約されていない

{1}サービス費用契約書締結
契約期間：{0}年1月1日～{0}年12月31日
契約金額：{4:,.2f}ドル
見積もり内容：
1.フォローサービス費:{2:,.2f}ドル
2.{0}年度ファイルメンテナンス費:{3:,.2f}ドル
3.その他フォローサービス関連費用:{5:,.2f}ドル
今後:毎年のサービス費用が変わるので、1回/年にて契約書を結ぶ必要あり

-------br-------br-------br-------br-------br-------

1、1月1日开始的合同，请说明事后申请的原因
-->并非事后申请，合同是每年签订的，但是要等每年单价出来后，才能确定合同，
   所以无法在{0}年1月1日之前提供下一年度的合同

2、说明涨价原因
-->UL GmbH每年会对费用进行调整。

-------br-------br-------br-------br-------br-------

合同金额(USD): {6:,.2f}
换算金额(基准货币): {7:,.2f}
        '''.format(contract_year, company_name, follow_cost, doc_cost, ttl_amt, other_cost, ttl_amt, ttl_amt_RMB))

        if self.txt_contract.get(0.0, END).strip() != '': self.txt_contract.delete(0.0, END)
        self.txt_contract.insert(INSERT, flowlites_content)


    def close_app(self):
        self.root.destroy()
        self.root.master.deiconify()

    def mainloop(self):
        self.root.mainloop()

def main():
    u = UL_Gmbh_contract()
    u.mainloop()

if __name__ == '__main__':
    main()
