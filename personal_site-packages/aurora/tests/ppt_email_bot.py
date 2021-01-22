from aurora.email.report_mail import ReportMail
from aurora.ppt.report_ppt import ReportPowerPoint
import os

if __name__ == "__main__":
    ppt_file = r"C:\Users\5106001995\Desktop\20210121 dummy\20201228 NX75_MP_Optical_Evaluation.pptx"
    img_dir  = r"C:\Users\5106001995\Desktop\20210121 dummy\slide_img"
    ppt = ReportPowerPoint(ppt_file, img_dir)
    ppt.shoot()
    ppt.quit()

    ppt_name = os.path.split(ppt_file)[-1].split('.')[0]
    dst_dir  = os.path.split(ppt_file)[0]
    rpt_mail = ReportMail(ppt_name, dst_dir, img_dir, ppt_file)
    rpt_mail.send()
    rpt_mail.quit()