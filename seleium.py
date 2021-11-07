from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from time import sleep
from PIL import Image
import sys
import requests
#字符串被当作url提交时会被自动进行url编码处理
from urllib.parse import quote
#quote() 明文转译文    unquote()译文转明文
import configparser
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from pytictoc import TicToc
from selenium.webdriver.firefox.service import Service
config = configparser.ConfigParser()
path = r'.\config.ini'
# 配置
print("Autor:冰美式不加糖\nNotice:此版本为1.0.0版本！\n本软件完全免费，如果你是付费购买，请立刻举报卖家！")
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
opt.add_argument("--headless")
# 禁用 gpu
opt.add_argument('--disable-gpu')
# 指定 firefox 的安装路径，如果配置了环境变量则不需指定
opt.binary_location = ".\Firefox\App\Firefox\\firefox.exe"
# 指定geckodriver的位置
s=Service(r'.\Firefox\geckodriver.exe')
browser = webdriver.Firefox( service=s, options=opt)
print("读取配置中…")
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
#登录webvpn
browser.find_element(By.ID,"user_login").send_keys(username)	
browser.find_element(By.ID,"user_password").send_keys(password1)	
browser.find_element(By.NAME,"commit").click()
print("请记住弹出的验证码,关闭后输入验证码,如果没有弹出验证码，请手动到目录里点开a.png")
logincheck=1
#开始验证验证码和大物实验平台的用户名和密码
while(logincheck):
    browser.find_element(By.ID,"user_account_numb").send_keys(username2)
    browser.find_element(By.ID,"user_password").send_keys(password2)
    #寻找验证码的图片，并对此截图保存并打开
    a = browser.find_element(By.CLASS_NAME,'rucaptcha-image')
    a.screenshot('a.png')
    img=Image.open('a.png')
    img.show()
    verify=input("请输入验证码:")
    #登入进大物平台
    browser.find_element(By.NAME,"_rucaptcha").send_keys(verify)
    browser.find_element(By.NAME,"commit").click()  
    sleep(1)#如果返回超时，可能导致程序崩溃。未作修改
    if(browser.find_element(By.XPATH,"//*[@id='flash-messages']/div").text=="验证码不正确，请重新输入！"):
        print("验证码输入错误，咱们能再来一次吗？")
    if(browser.find_element(By.XPATH,"//*[@id='flash-messages']/div").text=="登录成功！"):
        logincheck=0
    if(browser.find_element(By.XPATH,"//*[@id='flash-messages']/div").text=="用户名或密码错误，请重新输入！"):
        print("用户名或密码错误，请去config.ini重新调整大物平台的用户名或密码")
        browser.quit() 
# 获取当前标签页1的句柄
handle = browser.current_window_handle 
#读取课程，根据课程数新开标签页，正常情况会多开一个标签页，但并不影响，未作修改
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
print('想要选的课程总数 ' + str(len(all_handles)-1))
# 读入空余时间
freedays=[]
with open('freeday.txt') as f:
    for line in f:
        freedays = [int(i) for i in line.split( )]
        print (freedays)
        file.close()        
totaltime=0        
#这里是为了修复测试版本监控模式有些时候刷新超时
browser.set_page_load_timeout(3)
mode=1
#有些时候，控件没有加载出来导致程序崩溃，所以这里需要等待
wait2 = WebDriverWait(browser,10,0.5)
#监控模式代码
while(mode):
        #这里开始对每次遍历进行计时
        t = TicToc() 
        start_time= t.tic()
        #切换标签页，明确一点是切换每一个课程
        for page in range(len(all_handles)-1):
            cacheelem=wait2.until(lambda diver:str(browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr[1]/td[1]").text))
            browser.switch_to.window(browser.window_handles[page])
            wait2.until(lambda diver:browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[2]/span/b"))
            total=browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[2]/span/b")
            #对当前课程所有的星期和节次进行组合
            for i in range(int(total.text)):
                i=i+1
                weekday=browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr["+str(i)+"]/td[5]")
                part=browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr["+str(i)+"]/td[6]")
                combineday=int(weekday.text+part.text)
                #通过自由时间和组合的星期和节次进行匹配
                for p in range(len(freedays)):
                    #成功匹配成功
                    if(freedays[p]==combineday):
                        #开始对匹配成功的当前项进行计算是否有空余位置
                        if (str(browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr["+str(i)+"]/td[7]").text)>str(browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr["+str(i)+"]/td[8]").text)):
                            cacheelem=str(browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr[1]/td[1]").text)
                            browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr["+str(i)+"]/td[9]/a").click()
                            #点击后一定要等待返回的flashmesaage，否则下面获取不到信息会导致报错
                            wait2.until(lambda diver:browser.find_element(By.XPATH,"//*[@id='flash-messages']/div"))
                            #因为选课成功和本周以选完课都会返回第一个页面，所以只需要对一个在主页面不存在的元素进行捕获，如果捕获得到，则继续执行，否则返回抢课成功，并结束程序
                            try:
                             okla=str(browser.find_element(By.XPATH,"/html/body/main/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody[2]/tr[1]/td[1]").text)
                            except:
                              print("抢到课了"+cacheelem)
                              send_server("抢课成功通知:", cacheelem)  
                              #这里要确保完全退出
                              mode=0
                              sys.exit(0)
                            #这里开始对并没有出现上面的情况的条件进行解析    
                            cacheelem=okla
                            print(okla+str(combineday), end="")
                            wait2 = WebDriverWait(browser,10,0.5)
                            #这里需要等待返回，否则报错
                            wait2.until(lambda diver:browser.find_element(By.XPATH,"//*[@id='flash-messages']/div"))
                            if(browser.find_element(By.XPATH,"//*[@id='flash-messages']/div").text=="选课失败，该实验已满！"or"选课失败，所选实验时间无效！"):
                                print("课程无效或已满，下一个")
            print(cacheelem+"没有余量，继续监控")
            #这里开始对网页刷新进行捕获，有时的刷新并未成功，但并未停止，导致程序超时而崩溃
            while True:
                try:
                    browser.refresh()
                    break
                except:
                    pass 
        totaltime=totaltime+1    
        print("已经遍历次数",totaltime)     
        #输出此次遍历的时间                    
        end_time = t.toc()