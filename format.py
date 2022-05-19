from flask import Flask, request, jsonify
from datetime import timedelta
from gevent import pywsgi
import webbrowser, json, openpyxl, os, re

import flask

app = Flask(__name__, static_url_path="")
app.debug = False
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = timedelta(seconds=1)

HOST_PAGE = "http://localhost:40115"
HOST = "127.0.0.1"
PORT = 40115
VERSION = "v2.4.3"


class Excel_List:

    def __init__(self, path):  # 读取表格
        self.path = path.strip('"')
        self.sheet = openpyxl.load_workbook(self.path, data_only=True).active

    def is_correct_excel(self):  # 判断表格格式是否正确
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

    def return_col_data(self, col_num):  # 返回一列的数据并判断是否适合做关键字
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
        del col["delta"]

        #判断表格值是否为序号
        try:
            if sum([float(i) for i in col['values']
                    ]) < (self.last_line - self.first_line + 1)**2:
                col["isKeyWord"] = False
                col["reason"] = "该项可能是序号，不适合做关键字"
        except ValueError:
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


def read_json(path):
    current_path = os.getcwd()
    os.chdir(MY_PATH)
    with open(path, "r", encoding="utf-8-sig") as f:
        t = json.loads(f.read())
    os.chdir(current_path)
    return t


def write_json(path, obj):
    current_path = os.getcwd()
    os.chdir(MY_PATH)
    j = json.dumps(obj, ensure_ascii=False)
    with open(path, "w", encoding="utf-8-sig") as f:
        f.write(j)
    os.chdir(current_path)
    return j


@app.route("/", methods=["get"])
def index():
    if config["data"]:
        return app.send_static_file("index.html")
    else:
        return flask.redirect("/GetExcel")


@app.route("/GetExcel", methods=["get"])
def get_excel():
    return app.send_static_file("excel.html")


@app.route("/GetKeyWord", methods=["get"])
def get_key_word():
    return app.send_static_file("key.html")


@app.route("/FormatName", methods=["get"])
def format_name():
    return app.send_static_file("format_name.html")


@app.route("/Analysis", methods=["get"])
def analysis():
    return app.send_static_file("analysis.html")


@app.route("/log", methods=["get"])
def log():
    return app.send_static_file("update.html")


@app.route("/SubmitExcelPath", methods=["post"])
def submit_excel_path():
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


@app.route("/SubmitData", methods=["post"])
def submit_data():
    config["data"] = json.loads(request.data.decode())
    write_json("config.json", config)
    return jsonify({"code": 0, "msg": "上传成功"})


# 接收的参数为
# "path": 文件夹路径
# "execute": 命名列表[,,,'.xxx']
@app.route("/SubmitExecute", methods=["post"])
def submit_execute():
    a = json.loads(request.data.decode())
    a["path"] = os.path.normpath(a["path"].strip('"'))
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
    os.chdir(execute["path"])
    execute["old"] = [x for x in os.listdir() if os.path.isfile(x)]
    execute["data"] = [
        x for x in sorted(config["data"], key=lambda x: x["delta"])
        if x["isKeyWord"]
    ]
    return_new_name_list()
    find_new_name()


def return_new_name_list():
    new_name_list = []
    for i in range(len(config["data"][0]["values"])):
        name = ""
        for j in execute["execute"]:
            if type(j) == int:
                name += str(config["data"][j]["values"][i])
            elif type(j) == type(None):
                name += ""
            else:
                name += str(j)
        new_name_list += [name]
    execute["new"] = new_name_list


def find_new_name():
    execute["map"] = [0 for i in config["data"][0]["values"]]
    name_list = []
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
                                    "(%s)" % execute["map"][k] + last_name)
                    else:
                        new_name = execute["new"][k] + last_name
                    name_list += [{"old": execute["old"][i], "new": new_name}]
                    execute["map"][k] += 1
                    flag = 1  # 找到新名字以后，不需要匹配下一个关键字了
                    break
            if flag:
                break
    execute["list"] = name_list
    return name_list


def return_last_name(file_name):
    return re.findall(r"\.[^\.]+$", file_name)[0]


@app.route("/Rename", methods=["post"])
def Rename():
    if execute["flag"]:
        return jsonify({"code": 1, "msg": "已经改过名了，请勿重复点击"})
    else:
        new = [x["new"] for x in execute["list"]]
        if len(set(new)) < len(new):
            return jsonify({"code": 2, "msg": "新名字不能重复"})
        the_repeat_name = []
        for i in execute["list"]:
            try:
                os.rename(i["old"], i["new"])
            except FileExistsError:
                the_repeat_name.append(i)
        for i in the_repeat_name:
            os.rename(i["old"], i["new"])
        execute["flag"] = 1
        return jsonify({"code": 0, "msg": "改名成功"})


@app.route("/GetData", methods=["post"])
def get_data():
    return json.dumps(config["data"], ensure_ascii=False)


@app.route("/GetExecute", methods=["post"])
def get_execute():
    return jsonify({
        "map": execute["map"],
        "list": execute["list"],
        "new": execute["new"]
    })


@app.route("/GetVersion", methods=["post"])
def get_version():
    return jsonify({"version": VERSION})


@app.route("/show", methods=["get"])
def show():
    return json.dumps(config, ensure_ascii=False)


if __name__ == "__main__":
    MY_PATH = os.getcwd()
    try:
        config = read_json("config.json")
    except (FileNotFoundError, json.decoder.JSONDecodeError):
        config = {"data": []}
        write_json("config.json", config)
    execute = {}

    print("version:", VERSION)
    print(MY_PATH)
    webbrowser.open_new(HOST_PAGE)
    server = pywsgi.WSGIServer((HOST, PORT), app)
    server.serve_forever()
