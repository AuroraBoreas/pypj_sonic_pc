# M5 参照信号　自動調整　Program

import subprocess
import serial
import serial.tools.list_ports
import time
import openpyxl
import os

def search_secl_com_port():  ##Serial Consoleに有効なCOMポートを自動的に探して返す関数を定義
    coms = serial.tools.list_ports.comports()
    comlist = []
    for com in coms:
        comlist.append(com.device)
    print('\nConnected COM ports: ' + str(comlist)) ##COM PortのListを表示する
    use_port_secl = comlist[5]  ##Serial ConsoleのCOMポートがcomlist[0],[1],[2],[3]・・・の何番目か番号を格納
    print('\nUse COM port for Serial Console: ' + use_port_secl)  ##Serial ConsoleのCOMポート名を表示する
    
    return use_port_secl ##use_port_seclを戻り値として返す

def change_tv_volume(port, volume): #TV volume変更用関数
    ser = serial.Serial(port, 115200) #portと接続
    ser.reset_input_buffer() #input buffer削除
    ser.reset_output_buffer() #output buffer削除
    
    ser.write(b" \r\n")
    ser.write(b"su\r\n")
    ser.write(b"setenforce 0\r\n")
    ser.write(b"cli\r\n")
    ser.write(b"b.scm 0\r\n")
    ser.write(b"aud.uop.cv 0 0 " + volume.encode() + b"\r\n") # vol変更用命令文送信,int⇒Char型にしている
    ser.write(b"aud.uop.cv q\r\n")
    
    ser.close() #通信終了

def change_mic_gain(port): # mic gainを変更する関数
    ser = serial.Serial(port, 115200)  ##serというインスタンス(メインメモリ上に展開して処理・実行できる状態の実体)を作成している
    ser.reset_input_buffer() ##input bufferをreset
    ser.reset_output_buffer() ##output bufferをreset
    
    ser.write(b" \r\n")
    ser.write(b"su\r\n")
    ser.write(b"setenforce 0\r\n")
    ser.write(b"cli\r\n")
    ser.write(b"b.scm 0\r\n")
    ser.write(b"cm4.bf 4\r\n")
    print('\nShifted Mic gain to 4bit')  ##Micの4bit shiftが終了したことを表示する
    ser.close()  #通信終了
    
def change_src_gain(gain_value): #src_gain変更用関数
    print(subprocess.check_output("adb root", shell = True))
    subprocess.call("adb shell setenforce 0", shell = True)
    subprocess.call("adb shell chmod 777 data", shell = True)
    subprocess.call("adb shell setprop vendor.mtk.audio.aec.ref.gain "+ str(gain_value), shell = True) ##src gainの値を変更する
    print(str(gain_value))
    print("\nfin change_src_gain") ##sr gainの終了を表示
    
env = dict(os.environ) #環境変数取得
env["AUDIODEV"] = "1" #sox動作用に環境変数書き換え（TV:0の場合とTV:1の場合などがある）

now = time.ctime()
cnvtime = time.strptime(now)
print(time.strftime("\n%Y/%m/%d %H:%M", cnvtime))

vollist = [1,5]##,10,15,20,25,30,35,40,45,50,55,60,70,80,90,100] ##VolumeのListを定義する
res_list = []  ##結果を入れるlistを定義する

if __name__ == '__main__': #このmainから最初に動かす

    src_peak0_flag = 0 #src peak lv = 0になったことがあることを示するフラグ
    src_peak0_rms = 0 #はじめてsrc peak lv = 0になったときのrms
    fin_flag = 0 #最終ループを迎えたことを示すフラグ
    src_gain = 0 #src_gainの初期値
    mic_gain = 0 #src_gainの初期値
    fin_src_gain = 0 #最終的に書き込むsreg max値
    use_port_secl = search_secl_com_port()  ##この関数はSerial ConsoleのCOMポートをuse_port_seclとして返す
    change_mic_gain(port = use_port_secl) ##mic gainを変更する
    for number in vollist:
        
        if fin_flag == 1: #既に終了flagが立っている場合、fin_src_gainをレジスタに書き込んで次のループへ
            print("enter final processs...vol: ", number) #finフラグが立った後の一番最後の状態
            change_src_gain(gain_value = fin_src_gain ) #さちった時の処理
            
        else:             
            srcdiff = 0.0
            srcpeak = 100.0
            srcpeakrms = 0
            finalinput = 0
            
            #測定
            res = gain_value = 0 ##初期値の設定
            change_src_gain(gain_value) ##src gainを変更する
            print("\n***************************** Latest TV Volume = [ " + str(number) + " ] :  Gain_Value = [ " + str(res) + " ]   ************************************** ")

            change_tv_volume(port = use_port_secl, volume = str(number)) #TV volume変更用関数
            proc = subprocess.Popen(["sox", "C:\\Users\\Public\\Documents\\1_Auto_Evaluation\\whitenoise.wav", "-t","waveaudio", "0"], env=env) #whitenoiseを再生する  
#            proc2 = subprocess.Popen("C:\\Users\\Public\\Documents\\1_Auto_Evaluation\\srcadj_make25_for_Val_2.bat") #src,micのDataをdump
#            proc2.wait()
            proc2 = subprocess.Popen("C:\\Users\\Public\\Documents\\1_Auto_Evaluation\\srcadj_make25_for_Val_3.bat") #src,micのDataをdump
            proc2.wait()
            infile_path = r'/Users/Public/Documents/1_Auto_Evaluation/aec_loopback_post.bin' ##.bin fileのpathを通す
            outfile = open(r'/Users/Public/Documents/1_Auto_Evaluation/aec_loopback_post.raw', 'w') ##.rawfileをOpenする
            outfile.close()
            outfile = open(r'/Users/Public/Documents/1_Auto_Evaluation/aec_loopback_post.raw', 'ab') ##.raw fileをOpenする
            index = 0
            with open(infile_path, 'rb') as fp:  ##infileをOpenする
                while True:
                	buf = fp.read(2) ##bufferに2byte分入れる
                	if not buf: ##bufferにdataが無ければ
                		break
                	if(index%2 == 0): ##下位2byteを抽出
                		outfile.write(buf) ##outfileに書き込む
                	index += 1 ##indexに1を足す
            outfile.close() ##outfileを閉じる
            proc3 = subprocess.Popen("C:\\Users\\Public\\Documents\\1_Auto_Evaluation\\srcadj_make25_for_Val_4.bat") #src,micのDataを5秒間のwavにする
            proc3.wait()
            proc = subprocess.run(["sox", "C:\\Users\\Public\\Documents\\1_Auto_Evaluation\\src_25.wav", "-n", "stats"], stdout = subprocess.PIPE, stderr = subprocess.PIPE) #wav statsでpeak levelとRMS Levelを取得しstderrに目的の出力.
            
            preoutlog = proc.stderr.decode("utf8", "ignore") #stderrに目的の出力.proc.stderr.decodeという文字列をutf-8にする。 引数にignoreを入れてエラー回避。ignorを入れると表示できない文字を引っ張り出してくれる。
            outlog = preoutlog.splitlines() #出力を行ごとに分割
        
            Pk_lev_line = [line for line in outlog if "Pk lev dB" in line] #Pk lev dBが含まれる行抜き出し
            RMS_lev_line = [line for line in outlog if "RMS lev dB" in line] #RMS lev dBが含まれる行抜き出し
            
            print("\nsrc_Pk_lev= ")
            print(Pk_lev_line[0].split(" "))
            print("src_RMS_lev= ")
            print(RMS_lev_line[0].split(" "))
            
            src_Pk_lev = Pk_lev_line[0].split()[3] #Overall抜き出し・・・値だけ引っ張ってくる
            src_RMS_lev = RMS_lev_line[0].split()[3] #Overall抜き出し
            
            
            proc = subprocess.run(["sox", "C:\\Users\\Public\\Documents\\1_Auto_Evaluation\\mic_25.wav", "-n", "stats"], stdout = subprocess.PIPE, stderr = subprocess.PIPE)
            
            preoutlog = proc.stderr.decode("utf8") #stderr側に目的の出力
            outlog = preoutlog.splitlines() #出力を行ごとに分割
            
            Pk_lev_line = [line for line in outlog if "Pk lev dB" in line] #Pk lev dBが含まれる行抜き出し
            RMS_lev_line = [line for line in outlog if "RMS lev dB" in line] #RMS lev dBが含まれる行抜き出し
            

            print("mic_Pk_lev= ")
            print(Pk_lev_line[0].split(" "))
            print("mic_RMS_lev= ")
            print(RMS_lev_line[0].split(" "))
            
            mic_Pk_lev = Pk_lev_line[0].split()[3] #Overall抜き出し・・・値だけ引っ張ってくる
            mic_RMS_lev = RMS_lev_line[0].split()[3] #Overall抜き出し
    
            print("\nSRC_Peak_Level= " + src_Pk_lev)
            print("\nMIC_Peak_Level= " + mic_Pk_lev)
            
            print("\nSRC_RMS_Level= " + src_RMS_lev)
            print("\nMIC_RMS_Level= " + mic_RMS_lev)
            
            
            srcdiff = float(src_RMS_lev) - float(mic_RMS_lev)
            print("\nsrcdiff_Level= ", round(srcdiff,2)) #数字を四捨五入している。小数点第2が結果として出る。
            srcpeak = float(src_Pk_lev)
            
            while round(srcdiff,2) < 5.8 or 6.2 < round(srcdiff,2): #srcdiffが5.8から6.2の間実行。
                #書き込む値を計算
                input_src_gain = round(6.0 - round(srcdiff,2),2) #6.0からsrcdiffを引いた値をinput_src_gainに入れる
                print("\nAdjustment gain_value of difference = ",round(6.0 - round(srcdiff,2),2)) #TV Volume値とsrcdiffを6.0から引いた値を表示。
                if srcpeak == 0:  #src peak lev = 0になったとき
                    fin_flag = 1  #終了フラグを立てる
                    fin_src_gain = min([input_src_gain, 144]) #fin_src_gainの最大値は144以上になった時にfin_src_gainに144を入れる
                    change_src_gain(gain_value = fin_src_gain)
                    res=fin_src_gain
                    res_list.append(res)
                    print("\n****enable Fin flag because of src rms Lv + 3dB****")
                    break  #while srcdiff < 5.8 or 6.2 < srcdiff: のwhileを抜ける
                                            
                if input_src_gain > 144: #もし、input_src_gainが144より大きい時
                    fin_flag = 1
                    fin_src_gain = 144
                    change_src_gain(gain_value = fin_src_gain ) #src_gainに最後の値を書き込む
                    res=fin_src_gain
                    res_list.append(res)
                    print("\n****enable Fin flag because of input reg > 144****  gain_value = ", input_src_gain)
                    break  #while srcdiff < 5.8 or 6.2 < srcdiff: のwhileを抜ける
                    
                    
                change_src_gain(gain_value = input_src_gain) #input_src_gainに値を書き込む
                res=input_src_gain
                res_list.append(res)
                
                ###　Gain変更後のsrc_diffを測定
                print("\n***************************** TV Volume after adjustment = [ " + str(number) + " ] input_src_gain = [ " + str(res) + " ]  ************************************** ")
                
                change_tv_volume(port = use_port_secl, volume = str(number))
                proc = subprocess.Popen(["sox", "C:\\Users\\Public\\Documents\\1_Auto_Evaluation\\whitenoise.wav", "-t","waveaudio", "0"], env=env)  
#                proc2 = subprocess.Popen("C:\\Users\\Public\\Documents\\1_Auto_Evaluation\\srcadj_make25_for_Val_2.bat")
#                proc2.wait()
                proc2 = subprocess.Popen("C:\\Users\\Public\\Documents\\1_Auto_Evaluation\\srcadj_make25_for_Val_3.bat")
                proc2.wait()
                infile_path = r'/Users/Public/Documents/1_Auto_Evaluation/aec_loopback_post.bin'
                outfile = open(r'/Users/Public/Documents/1_Auto_Evaluation/aec_loopback_post.raw', 'w')
                outfile.close()
                outfile = open(r'/Users/Public/Documents/1_Auto_Evaluation/aec_loopback_post.raw', 'ab')
                index = 0
                with open(infile_path, 'rb') as fp:
                	while True:
                		buf = fp.read(2)
                		if not buf:
                			break
                		if(index%2 == 0):
                			outfile.write(buf)
                		index += 1
                outfile.close()
                proc3 = subprocess.Popen("C:\\Users\\Public\\Documents\\1_Auto_Evaluation\\srcadj_make25_for_Val_4.bat") #src,micのDataを5秒間のwavにする
                proc3.wait()
                proc = subprocess.run(["sox", "C:\\Users\\Public\\Documents\\1_Auto_Evaluation\\src_25.wav", "-n", "stats"], stdout = subprocess.PIPE, stderr = subprocess.PIPE)
                
                preoutlog = proc.stderr.decode("utf8", "ignore") #stderrに目的の出力. ignoreでエラー回避
                outlog = preoutlog.splitlines() #出力を行ごとに分割
            
                Pk_lev_line = [line for line in outlog if "Pk lev dB" in line] #Pk lev dBが含まれる行抜き出し
                RMS_lev_line = [line for line in outlog if "RMS lev dB" in line] #RMS lev dBが含まれる行抜き出し
                
                print("\nsrc_Pk_lev= ")
                print(Pk_lev_line[0].split(" "))
                print("src_RMS_lev= ")
                print(RMS_lev_line[0].split(" "))
                
                src_Pk_lev = Pk_lev_line[0].split()[3] #Overall抜き出し
                src_RMS_lev = RMS_lev_line[0].split()[3] #Overall抜き出し
                
                
                proc = subprocess.run(["sox", "C:\\Users\\Public\\Documents\\1_Auto_Evaluation\\mic_25.wav", "-n", "stats"], stdout = subprocess.PIPE, stderr = subprocess.PIPE)
                
                preoutlog = proc.stderr.decode("utf8") #stderr側に目的の出力
                outlog = preoutlog.splitlines() #出力を行ごとに分割
                
                Pk_lev_line = [line for line in outlog if "Pk lev dB" in line] #Pk lev dBが含まれる行抜き出し
                RMS_lev_line = [line for line in outlog if "RMS lev dB" in line] #RMS lev dBが含まれる行抜き出し
                
                print("mic_Pk_lev= ")
                print(Pk_lev_line[0].split(" "))
                print("mic_RMS_lev= ")
                print(RMS_lev_line[0].split(" "))
                
                mic_Pk_lev = Pk_lev_line[0].split()[3] #Overall抜き出し
                mic_RMS_lev = RMS_lev_line[0].split()[3] #Overall抜き出し
        
                print("\nSRC_Peak_Level= " + src_Pk_lev)
                print("\nMIC_Peak_Level= " + mic_Pk_lev)
                
                print("\nSRC_RMS_Level= " + src_RMS_lev)
                print("\nMIC_RMS_Level= " + mic_RMS_lev)
                
                srcdiff = float(src_RMS_lev) - float(mic_RMS_lev)
                print("\nsrcdiff_Level= ", round(srcdiff,2)) #数字を四捨五入している。小数点第2まででるので小数点第1が結果として出る。
                srcpeak = float(src_Pk_lev)
                
                time.sleep(5)
    wb = openpyxl.load_workbook("C:\\Users\\Public\\Documents\\1_Auto_Evaluation\\Auto_Evaluation_Excel_SRC_Gain_Value1.xlsx")        
    sheet = wb["Auto_Evaluation_Excel_Data"] #sheetを指定する
    
    index = 0
    for line, number in enumerate(vollist): #for文をvolume listの個数分回している
        print("\nTV Volume [ " + str(number) + " ] : final src_gain = [ " + str(res_list) + " ]")
        print(line)
        print(res_list)
        sheet.cell(row = 3 + line, column = 2).value = res_list[index] #3行2列目にｓｒｃ参照信号の値を書き込む
        index += 1
    wb.save("C:\\Users\\Public\\Documents\\1_Auto_Evaluation\\Auto_Evaluation_Excel_SRC_Gain_Value2.xlsx") #指定のExcel fileをsaveする
   
print("****************************************  Finish  *********************************************************** ")