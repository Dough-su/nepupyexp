from os import system
from requests.api import options
from selenium import webdriver
from selenium.webdriver.firefox import service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from time import sleep, strftime, localtime, time
from PIL import Image
import requests
#字符串被当作url提交时会被自动进行url编码处理
from urllib.parse import quote,unquote
#quote() 明文转译文    unquote()译文转明文
import configparser
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from pytictoc import TicToc
from selenium.webdriver.firefox.service import Service
config = configparser.ConfigParser()
path = r'.\config.ini'
# 配置
print("Autor:冰美式不加糖\nNotice:此版本为预发布版本！\n本软件完全免费，如果你是付费购买，请立刻举报卖家！")
config.read(path)
notify=config['notify']['serverchan']
def send_server(receiver, text):
    api = "https://sctapi.ftqq.com/"+str(notify)+".send" 
    data = {
            'text':receiver, #标题
            'desp':text} #内容
    result = requests.post(api, data = data)
    return(result)
opt = webdriver.FirefoxOptions()
# 设置无界面
#opt.add_argument("--headless")
# 禁用 gpu
opt.add_argument('--disable-gpu')
# 指定 firefox 的安装路径，如果配置了环境变量则不需指定
opt.binary_location = ".\Firefox\App\Firefox\\firefox.exe"
s=Service(r'.\Firefox\geckodriver.exe')
browser = webdriver.Firefox( service=s, options=opt)
print("读取配置中…")
#抢课模式，监控模式，监控模式为1，抢课模式为0,默认为1
mode=config['mode']['mode']
mode=int(mode)
#webvpn的学号和密码
username=config['webvpn']['username']
password1=config['webvpn']['password']
#大物平台的密码
username2=config['experiment']['username']
password2=config['experiment']['password']
#上学期还是下学期，上学期填1，下学期填2
half_semi=config['semi']['semi']
#今年是哪年，如果是2021年，那么填写2021-2022
semi_name=config['semi']['year']
#想要抢第几周的大物课第几周就填几
week=config['semi']['week']
#通过浏览器向服务器发送URL请求
browser.get("https://58-155-46-47.webvpn.nepu.edu.cn/")
#不要设置浏览器大小，设置大小后面无法输入文本
print("开始登陆中…")
browser.find_element(By.ID,"user_login").send_keys(username)	
browser.find_element(By.ID,"user_password").send_keys(password1)	
browser.find_element(By.NAME,"commit").click()
print("请记住弹出的验证码,关闭后输入验证码,如果没有弹出验证码，请手动到目录里点开a.png")
#得到验证码在屏幕中的坐标位置
logincheck=1
while(logincheck):
    browser.find_element(By.ID,"user_account_numb").send_keys(username2)
    browser.find_element(By.ID,"user_password").send_keys(password2)
    a = browser.find_element(By.CLASS_NAME,'rucaptcha-image')
    a.screenshot('a.png')
    img=Image.open('a.png')
    img.show()
    verify=input("请输入验证码:")
    #登入进大物平台
    browser.find_element(By.NAME,"_rucaptcha").send_keys(verify)
    browser.find_element(By.NAME,"commit").click()  
    sleep(2)
    if(browser.find_element(By.XPATH,"//*[@id='flash-messages']/div").text=="验证码不正确，请重新输入！"):
        print("验证码输入错误，咱们能再来一次吗？")
    if(browser.find_element(By.XPATH,"//*[@id='flash-messages']/div").text=="登录成功！"):
        logincheck=0
    if(browser.find_element(By.XPATH,"//*[@id='flash-messages']/div").text=="用户名或密码错误，请重新输入！"):
        print("用户名或密码错误，请去config.ini重新调整大物平台的用户名或密码")
        browser.quit() 
# 获取当前标签页1的句柄
handle = browser.current_window_handle 
#读取课程
# 根据课程数新开标签页
i=0
file=open("courses.txt", encoding="utf-8")
for line in file:
    line = line.split()
    browser.execute_script('window.open("","_blank");')
    browser.switch_to.window(browser.window_handles[i])
    dataset = str(line).replace("[", "")
    dataset = dataset.replace("]", "")
    dataset=dataset.replace(",","")
    dataset=dataset.replace("'","")
    browser.get('https://58-155-46-47.webvpn.nepu.edu.cn/experiments?exp_name='+quote(str(dataset))+'&half_semi='+half_semi+'&semi_name='+semi_name+'&week='+week+'')
    i=i+1
file.close()
# 获取所有的打开的标签页句柄
all_handles = browser.window_handles
#这里要减去第一个负责登陆的页面
print('课程总数 ' + str(len(all_handles)-1))
# 切换到标签页1
freedays=[]
with open('freeday.txt') as f:
    for line in f:
        freedays = [int(i) for i in line.split( )]
        print (freedays)
        file.close()
        
totaltime=0        
browser.set_page_load_timeout(3)#这里是为了修复测试版本有时网页刷新超时 
while(mode):
    t = TicToc()  # create instance of class
    start_time= t.tic() #start timer
    for page in range(len(all_handles)-1):
        browser.switch_to.window(browser.window_handles[page])
        wait = WebDriverWait(browser,10,0.5)
        wait.until(lambda diver:browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[2]/span/b"))
        total=browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[2]/span/b")
        for i in range(int(total.text)):
            i=i+1
            weekday=browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr["+str(i)+"]/td[5]")
            part=browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr["+str(i)+"]/td[6]")
            combineday=int(weekday.text+part.text)
            for p in range(len(freedays)):
                if(freedays[p]==combineday):
                    browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr["+str(i)+"]/td[9]/a").click()
                    okla=str(browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr[1]/td[1]").text)+str(combineday)
                    print("已选中"+str(browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr[1]/td[1]").text)+str(combineday), end="")
                    wait2 = WebDriverWait(browser,10,0.5)
                    wait2.until(lambda diver:browser.find_element(By.XPATH,"//*[@id='flash-messages']/div"))
                    if(browser.find_element(By.XPATH,"//*[@id='flash-messages']/div").text=="选课失败，该实验已满！"or"选课失败，所选实验时间无效！"):
                        print("此课满课,下一个")
                    else:
                        print("选课应该或许成功吧！")
                        send_server("抢课成功通知:", okla)
                        mode=0
    totaltime=totaltime+1    
    print("已经遍历次数",totaltime)                
    end_time = t.toc()
mode=mode^1
while(mode):
    try:
        t = TicToc()  # create instance of class
        start_time= t.tic() #start timer
        for page in range(len(all_handles)-1):
            browser.switch_to.window(browser.window_handles[page])
            wait = WebDriverWait(browser,10,0.5)
            wait.until(lambda diver:browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[2]/span/b"))
            total=browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[2]/span/b")
            for i in range(int(total.text)):
                i=i+1
                weekday=browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr["+str(i)+"]/td[5]")
                part=browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr["+str(i)+"]/td[6]")
                combineday=int(weekday.text+part.text)
                for p in range(len(freedays)):
                    if(freedays[p]==combineday):
                        if (str(browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr["+str(i)+"]/td[7]").text)>str(browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr["+str(i)+"]/td[8]").text)):
                            browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr["+str(i)+"]/td[9]/a").click()
                            okla=str(browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr[1]/td[1]").text)+str(combineday)
                            print("已选中"+str(browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr[1]/td[1]").text)+str(combineday), end="")
                            wait2 = WebDriverWait(browser,10,0.5)
                            wait2.until(lambda diver:browser.find_element(By.XPATH,"//*[@id='flash-messages']/div"))
                            if(browser.find_element(By.XPATH,"//*[@id='flash-messages']/div").text=="选课失败，该实验已满！"or"选课失败，所选实验时间无效！"):
                                print("课程无效或已满，下一个")
                            else:
                                print("选课应该或许成功吧！")
                                send_server("抢课成功通知:", okla)  
                                mode=0
            print("当前时间"+str(browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr[1]/td[1]").text)+"没有余量，继续监控")
            while True:
                try:
                    browser.refresh()
                    break
                except:
                    print("捕获到刷新超时，以重试(下一个版本此提示取消)")
                    send_server("超时通知:", "就是单纯超时了")  
                    pass 
    except Exception as msg:
    # 时间戳名称，防止覆盖
        name = time.strftime("%H.%M.%S")
        # 异常截图保存在本地
        browser.get_screenshot_as_file('%s.png'%name)

    totaltime=totaltime+1    
    print("已经遍历次数",totaltime)                         
    end_time = t.toc()
browser.quit    
