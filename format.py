from flask import Flask, request, jsonify, redirect
from datetime import timedelta
from gevent import pywsgi
import webbrowser, json, openpyxl, os, re, time, shutil

app = Flask(__name__, static_url_path="")
app.debug = False
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = timedelta(seconds=1)

HOST_PAGE = "http://localhost:40115"
HOST = "127.0.0.1"
PORT = 40115
VERSION = "v2.6.0"


class Excel_List:

    def __init__(self, path):  # 读取表格
        self.path = path.strip('"')
        self.sheet = openpyxl.load_workbook(self.path, data_only=True).active

    def is_correct_excel(self):
        '''
        判断表格格式是否正确
        
        返回值：
        
        1：表格正确

        0：表格为空

        -1：表格没超过2行

        2：关键字有重复
        '''
        if self.sheet.cell(1, 1).value == None:
            return 0
        if self.sheet.max_column <= 1:  # 单列表格
            self.first_line = 1
            self.find_last_line()
            if self.last_line < 2:
                return -1
            self.first_col = 1
            self.last_col = 1
            return 1
        elif self.sheet.max_column > 1:  # 多列表格
            for i in range(1, 10):  # 判断是否存在表格标题
                if self.sheet.cell(i, 2).value != None:
                    self.first_line = i
                    self.find_last_line()
                    if self.last_line < 2:
                        return -1
                    self.first_col = 1
                    self.find_last_col()
                    key_list = [
                        self.sheet.cell(self.first_line, i).value
                        for i in range(self.first_col, self.last_col + 1)
                    ]
                    if len(set(key_list)) < len(key_list):
                        return 2
                    return 1

    def find_last_line(self):
        i = self.first_line
        while 1:
            if self.sheet.cell(i, 1).value == None:
                break
            i += 1
        self.last_line = i - 1

    def find_last_col(self):
        i = 1
        while 1:
            if self.sheet.cell(self.first_line, i).value == None:
                break
            i += 1
        self.last_col = i - 1

    def return_col_data(self, col_num: int) -> dict:
        ''' 
        返回一列的数据并判断是否适合做关键字

        ['key']关键字

        ['value']值列表

        ['isKeyWord']是否为关键字

        ['reason']可作为关键字的原因

        ['delta']该关键字的重复次数，用于给关键字排序
        '''
        col = {}
        col["key"] = self.sheet.cell(self.first_line,
                                     col_num).value  # 关键字是该列的第一行
        col["values"] = [
            self.sheet.cell(i, col_num).value
            for i in range(self.first_line + 1, self.last_line + 1)
        ]  # 表里的值是表格除了标题以外的东西

        #判断表格值是否包含重复项
        col["delta"] = len(col["values"]) - len(set(col["values"]))
        if col["delta"] > 2:
            col["isKeyWord"] = False
            col["reason"] = "该项不适合做关键字，因为至少有%d个重复项" % col["delta"]
        else:
            col["isKeyWord"] = True
            col["reason"] = "该项适合做关键字"

        #判断表格值是否为序号
        try:
            if sum([float(i) for i in col['values']
                    ]) < (self.last_line - self.first_line + 1)**2:
                col["isKeyWord"] = False
                col["reason"] = "该项可能是序号，不适合做关键字"
        except (ValueError, TypeError):
            pass
        return col

    def return_excel_data(self):  # 构成关键字表格
        return [
            self.return_col_data(x)
            for x in range(self.first_col, self.last_col + 1)
        ]

    def show_excel(self):
        sheet = self.return_excel_data()
        for i in sheet:
            print(i["key"], end="\t")
        print()
        for j in range(self.last_line - self.first_line):
            for i in sheet:
                print(i["values"][j], end="\t")
            print()
        print("共%d行" % j + 1)


def read_json(path: str):
    current_path = os.getcwd()
    os.chdir(MY_PATH)
    with open(path, "r", encoding="utf-8-sig") as f:
        t = json.loads(f.read())
    os.chdir(current_path)
    return t


def write_json(path: str, obj):
    current_path = os.getcwd()  # 保存当前路径
    os.chdir(MY_PATH)  # 切换回程序所在的路径写入配置文件
    j = json.dumps(obj, ensure_ascii=False)
    with open(path, "w", encoding="utf-8-sig") as f:
        f.write(j)
    os.chdir(current_path)
    return j


@app.route("/", methods=["get"])
def index():
    '''
    根目录

    有数据直接进入主页

    没有数据则选择excel
    '''
    if config["data"]:
        return app.send_static_file("index.html")
    else:
        return redirect("/GetExcel")


@app.route("/GetExcel", methods=["get"])
def get_excel():
    '''
    获取excel
    '''
    return app.send_static_file("excel.html")


@app.route("/GetKeyWord", methods=["get"])
def get_key_word():
    '''
    获取关键字
    '''
    return app.send_static_file("key.html")


@app.route("/Analysis", methods=["get"])
def analysis():
    '''
    数据分析页面
    '''
    return app.send_static_file("analysis.html")


@app.route("/log", methods=["get"])
def log():
    '''
    版本更新日志页面
    '''
    return app.send_static_file("update.html")


@app.route("/SubmitExcelPath", methods=["post"])
def submit_excel_path():
    '''
    提交excel的路径接口

    接收：

    path：excel的路径
    
    返回值：

    0：读取成功

    1：文件不存在

    2：文件扩展名不对

    3：文件为空

    4：文件少于2行
    '''
    path = os.path.normpath(
        json.loads(request.data.decode())["path"].strip('"'))
    if not os.path.exists(path):
        return jsonify({"code": 1, "msg": "提交失败，文件不存在"})
    if not re.search(
            '^[A-Za-z]:\\\\([^|><?*":\\/]*\\\\)*[^|><?*":\\/]*\.((xlsx)|(xlsm)|(xltx)|(xltm))$',
            path,
    ):
        return jsonify({
            "code":
            2,
            "msg":
            "提交失败，不能读取该格式文件，请选择(.xlsx)(.xlsm)(.xltx)(.xltm)文件"
        })
    sheet = Excel_List(path)
    code = sheet.is_correct_excel()
    if code == 1:
        config["data"] = sheet.return_excel_data()
        return json.dumps({"code": 0, "msg": "读取成功"}, ensure_ascii=False)
    elif code == 0:
        return json.dumps({"code": 3, "msg": "文件不能为空"}, ensure_ascii=False)
    elif code == -1:
        return json.dumps({"code": 4, "msg": "文件应至少有2行"}, ensure_ascii=False)
    elif code == 2:
        return json.dumps({"code": 5, "msg": "文件表头存在重复"}, ensure_ascii=False)


@app.route("/SubmitData", methods=["post"])
def submit_data():
    '''
    提交关键字配置

    接收：

    关键值的数据

    返回值：

    0：数据保存成功
    '''
    config["data"] = json.loads(request.data.decode())
    write_json("config.json", config)
    return jsonify({"code": 0, "msg": "上传成功"})


@app.route("/SubmitExecute", methods=["post"])
def submit_execute():
    '''
    接收要改名的文件夹路径
    
    接收：
    
    path：文件夹路径

    execute：命名列表[,,,'.xxx']

    返回值：

    0：文件夹获取成功

    1：路径不存在

    2：路径为文件路径

    3：路径包含了自身程序

    4：路径格式不对
    
    '''
    a = json.loads(request.data.decode())
    a["path"] = os.path.normpath(a["path"].strip('"'))  # 格式化路径为标准格式
    if not re.search('^[A-Za-z]:\\\\([^|><?*":\\/]*\\\\)*([^|><?*":\\/]*)?$',
                     a["path"]):
        return jsonify({"code": 4, "msg": "提交失败，不是路径的标准格式"})
    if a["path"] in MY_PATH:
        return jsonify({"code": 3, "msg": "提交失败，请不要提交包含本程序的路径"})
    if os.path.exists(a["path"]):
        if os.path.isdir(a["path"]):
            execute.update(a)
            execute["flag"] = 0
            return_old_and_new_name_compare()
            return jsonify({"code": 0, "msg": "提交成功"})
        else:
            return jsonify({"code": 2, "msg": "提交失败，请提交一个目录而非文件"})
    else:
        return jsonify({"code": 1, "msg": "提交失败，目录不存在"})


def return_old_and_new_name_compare():
    os.chdir(execute["path"])  # 切换到文件夹目录
    execute["old"] = [x for x in os.listdir()
                      if os.path.isfile(x)]  # 获取文件夹里文件的名称
    execute["data"] = [
        x for x in sorted(config["data"], key=lambda x: x["delta"])
        if x["isKeyWord"]
    ]  #为关键字进行排序
    return_new_name_list()  # 生成一个新名字列表
    find_new_name()


def return_new_name_list():
    new_name_list = []
    for i in range(len(config["data"][0]["values"])):
        name = ""
        for j in execute["execute"]:
            if type(j) == int:
                temp = config["data"][j]["values"][i]
                if temp:  # 如果该项在excel里为空值则标明
                    name += str(temp)
                else:
                    name += str("空值")
            elif type(j) == type(None):
                name += ""
            else:
                name += str(j)
        new_name_list += [name]
    execute["new"] = new_name_list


def find_new_name():
    execute["map"] = [0 for _ in config["data"][0]["values"]]  # 生成一个匹配成功次数列表
    name_list = []  # 新旧名称对照表
    for i in range(len(execute["old"])):  # 遍历旧名字，i为旧名字的序数
        flag = 0
        last_name = return_last_name(execute["old"][i])
        for j in execute["data"]:  # 遍历所有关键字，j为关键字的键值对
            for k in range(len(j["values"])):  # 遍历该关键字的值列表，k为值的序数
                # if not execute['map'][k]:#只要匹配到了一次就不匹配下一个关键字
                if re.search(str(j["values"][k]),
                             execute["old"][i]):  # 在旧名字里匹配关键字
                    if execute["map"][k] > 0:
                        new_name = (execute["new"][k] +
                                    "(%s)" % execute["map"][k] + last_name
                                    )  #如果多次匹配成功就在新名字后面增加序号
                    else:
                        new_name = execute["new"][
                            k] + last_name  #第一次匹配成功就确定一个新名字
                    name_list += [{"old": execute["old"][i], "new": new_name}]
                    execute["map"][k] += 1
                    flag = 1  # 为这个旧名字找到新名字以后，不需要匹配下一个关键字了
                    break
            if flag:
                break
    execute["list"] = name_list


def return_last_name(file_name):
    '''
    确定后缀名
    '''
    return re.findall(r"\.[^\.]+$", file_name)[0]


@app.route("/Rename", methods=["post"])
def Rename():
    '''
    发起重命名的请求
    
    返回值：
    
    0：改名成功
    
    1：重复点击改名
    
    2：新名字不能重复
    '''
    if execute["flag"]:
        return jsonify({"code": 1, "msg": "已经改过名了，请勿重复点击"})
    else:
        new = [_["new"] for _ in execute["list"]]
        if len(set(new)) < len(new):
            return jsonify({"code": 2, "msg": "新名字不能重复"})
        the_repeat_name = []
        for i in execute["list"]:
            try:
                os.rename(i["old"], i["new"])
            except FileExistsError:  # 有的旧名字肯和有的新名字一样
                the_repeat_name.append(i)
        for i in the_repeat_name:
            os.rename(i["old"], i["new"])
        execute["flag"] = 1
        return jsonify({"code": 0, "msg": "改名成功"})


@app.route("/Recover", methods=["post"])
def Recover():
    '''
    发起恢复的请求
    
    返回值：
    
    0：恢复成功

    1: 尚未进行重命名
    '''
    if execute["flag"]:
        the_repeat_name = []
        for i in execute["list"]:
            try:
                os.rename(i["new"], i["old"])
            except FileExistsError:
                the_repeat_name.append(i)
        for i in the_repeat_name:
            os.rename(i["old"], i["new"])
        execute["flag"] = 0
        return jsonify({"code": 0, "msg": "恢复成功"})
    else:
        return jsonify({"code": 1, "msg": "尚未进行重命名"})


@app.route("/Backup", methods=["post"])
def Backup():
    '''
    发起恢复的请求
    
    返回值：
    
    0：恢复成功

    1: 已经改过名字了

    2：目录已存在
    '''
    if execute["flag"]:
        return jsonify({"code": 1, "msg": "已经改名请先恢复再备份"})
    else:
        t = time.time()
        name = time.strftime("%Y年%m月%d日 %H时%M分%S",
                             time.localtime(t)) + str(t - int(t))[1:5] + "秒 备份"
        try:
            os.mkdir(name)
        except FileExistsError:
            return jsonify({"code": 2, "msg": "目录已存在，请重试"})
        for i in execute["old"]:
            shutil.copy(i, name)
        return jsonify({"code": 0, "msg": "备份成功"})


@app.route("/GetData", methods=["post"])
def get_data():
    '''
    获取当前程序的数据

    返回值：

    数据列表
    '''
    return json.dumps(config["data"], ensure_ascii=False)


@app.route("/GetExecute", methods=["post"])
def get_execute():
    '''
    获取本次改名的操作以及新旧名字

    返回值：

    "map"：匹配成功的次数标记

    "list"：新旧名字对应表

    "new"：新名字

    "old"：旧名字
    '''
    return jsonify({
        "map": execute["map"],
        "list": execute["list"],
        "new": execute["new"],
        "old": execute["old"],
        "flag": execute["flag"]
    })


@app.route("/GetVersion", methods=["get", "post"])
def get_version():
    '''
    获取版本信息
    
    返回值：
    
    版本号'''
    return jsonify({"version": VERSION})


if __name__ == "__main__":
    MY_PATH = os.getcwd()
    try:
        config = read_json("config.json")
    except (FileNotFoundError, json.decoder.JSONDecodeError):
        config = {"data": []}
        write_json("config.json", config)
    execute = {"flag": 0}
    # execute各项解释
    # flag：是否使用过，1为使用过，0为未使用
    # path：要操作的目录
    # data：关键词排序过后的数据
    # execute：命名的格式【关键字序号，分隔符，关键字序号....】
    # new：新名字列表
    # map：名字匹配成功的次数标记
    # list：新旧名字对照表

    print("version:", VERSION)
    print(MY_PATH)
    webbrowser.open_new(HOST_PAGE)
    server = pywsgi.WSGIServer((HOST, PORT), app)
    server.serve_forever()
