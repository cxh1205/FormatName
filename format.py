#! python3
#coding=utf-8
import win32api, openpyxl, os, re, traceback


BASIC = 'Class Tool v1.0'
file_name = '配置文件.jitang'


def set():
    excel_dir = input('请输入数据列表位置：')
    flag_1 = input_int('是否输入基础信息？')
    keys, values = [],[]
    while True:
        if flag_1:
            keys += [input('请输入基础信息名称：')]
            values += [input('请输入基础信息的值：')]
        else:
            break
        flag_1 = input_int('是否继续输入基础信息？')
    setting = [BASIC,  '0,,0,,0,,0', ','.join(keys), ','.join(values)] + [','.join(i) for i in read_excel(excel_dir)]
    save_setting(setting)
    print('配置信息录入完成\n')


def save_setting(setting, end = '\n'):
    with open(file_name, 'w', encoding = 'utf-8') as f:
            for line in setting:
               f.write(line + end)
    win32api.SetFileAttributes('配置文件.jitang', 6)

                
def read_excel(file_dir):
    file_dir = file_dir.strip('"').strip("'")
    sheet = openpyxl.load_workbook(file_dir).active
    first_line = 1
    while True:
        if sheet.cell(first_line, 2).value == None:
            first_line += 1
        else:
            break
    n = 1
    while True:
        if sheet.cell(first_line, n).value == None:
            n -= 1
            break
        n += 1
    excel_list = []
    for col_num in range(1, n+1):
        excel_list += [[]]
        row_num = first_line
        while True:
            if sheet.cell(row_num, col_num).value == None:
                break
            excel_list[col_num - 1] += [str(sheet.cell(row_num, col_num).value)]
            row_num += 1
    return excel_list
            
        
def read_setting():
    with open(file_name, 'r' ,encoding = 'utf-8') as f:
        setting = f.readlines()[1:]
    setting = [i.rstrip('\n').split(',') for i in setting]
    last_choice = setting[0]
    for i in range(4):
        last_choice[2*i] = int(last_choice[2*i])
    keys = [''] + [i for i in setting[1]] + [setting[i][0] for i in range(3, len(setting))]
    keys_num = len(setting[1])
    values = [i for i in setting[2]]
    info_list = [[setting[i][j] for i in range(3, len(setting))] for j in range(1, len(setting[3]))]
    name = []
    for i in range(4):
        name += keys[last_choice[2*i]]
        if i<3:
            name += last_choice[2*i+1]
    if name:
        print('上次的命名格式为：' + ''.join(name) + '.xxx')
    if setting[1][0]:
        print('现在拥有的选项为：' + ','.join(keys[1:]))
    else:
        print('现在拥有的选项为：' + ','.join(keys[2:]))
        keys_num = 0
        del keys[1]
    return {'last_choice':last_choice, 'keys':keys, 'keys_num':keys_num, 'values':values, 'info':info_list}
        

def input_int(question):
    while True:
        try:
            num = int(input(question + '是输1，否输0：'))
            if num == 1 or num == 0:
                break
            else:
                print('只能输入0或1！', end = ' ')
        except:
            print('请输入0或1！', end = ' ')
    return num


def start():
    while True:
        try:
            my_dir = os.getcwd()
            with open(file_name, 'r', encoding = 'utf-8') as f:
                if f.readlines()[0].rstrip('\n') != BASIC:#配置文件版本不对就重新设置
                    os.remove(file_name)
                    1/0
            dic = read_setting()
            flag = input_int('是否需要重新配置信息？')
            if flag:
                1/0
            flag1 = 0
            for i in dic['last_choice']:
                if i:
                    flag1 = 1
                    break
            if flag1:
                flag = input_int('是否使用上次命名格式？')
            else:
                flag = 0
            if flag:
                pass
            else:
                while True:
                    flag1 = input_int('是否统一分隔符？')
                    if flag1:
                        split = input('分隔符统一为：')
                        dic['last_choice'][1] = dic['last_choice'][3] = dic['last_choice'][5] = split
                    print('\n接下来将根据以下列表输入' + str(7-flag1*3) + '次信息')
                    print('不输入请按0')
                    for i in range(1, len(dic['keys'])):
                        print('输入 %s 请按%d'%(dic['keys'][i],i))
                    print()
                    for i in range(4):
                        dic['last_choice'][2*i] = int(input('请输入序号：'))
                        if not flag1 and i<3:
                            dic['last_choice'][2*i+1] = input('请输入分隔符：')
                    for i in range(1, 4):
                        if dic['last_choice'][2*i] == 0:
                            dic['last_choice'][2*i-1] = ''
                    name = []
                    for i in range(4):
                        name += dic['keys'][dic['last_choice'][2*i]]
                        if i<3:
                            name += dic['last_choice'][2*i+1]
                    print('命名格式为：' + ''.join(name) + '.xxx')
                    flag2 = input_int('对命名格式是否满意？')
                    if flag2:
                        with open(file_name, 'r', encoding = 'utf-8') as f:
                            setting = f.readlines()
                        last_choice = dic['last_choice'][0:]
                        for i in range(4):
                            last_choice[2*i] = str(last_choice[2*i])
                        setting[1] = ','.join(last_choice) + '\n'
                        save_setting(setting, end = '')
                        break
            file_dir = input('请输入要重新命名的文件夹路径：').strip('"').strip("'")
            main(file_dir, dic)
            print('重命名完成！\n')
            flag = input_int('是否继续使用？')
            if flag:
                os.chdir(my_dir)
                print()
            else:
                break
        except ZeroDivisionError:
            set()
        except FileNotFoundError:
            set()
        except Exception as e:
            print('\n' + '='*40)
            traceback.print_exc()
            print('='*40)
            input('↑↑↑↑↑请截图上面的错误反馈给我，谢谢↑↑↑↑↑\n截图后请复制进入https://class-tool.jitang.xyz添加我微信，向我发送截图\n\n如需继续使用请更改配置或更改重命名的文件夹路径\n请按回车继续……')
            print()
        


def main(file_dir, dic):
    map = [0 for i in range(len(dic['info']))]
    os.chdir(r'%s'% file_dir)
    dir_list = os.listdir()
    last_choice = [dic['last_choice'][2*i] for i in range(4)]
    for file_name in dir_list:
        for i in range(len(dic['info'])):
            for j in last_choice:
                if j > dic['keys_num']:
                    if re.search(r'%s'%dic['info'][i][j-1-dic['keys_num']], file_name):
                        try:
                            last_name = re.findall(r'\.[^\.]+$', file_name)[0]
                            map[i] += 1
                            if map[i] >= 2:
                                last_name = '(%d)'%map[i] + last_name
                            os.rename(file_name, name(dic, i) + last_name)
                        except:
                            pass
                        break
    for i in range(len(map)):
        if map[i] > 1:
            print('%s 交重复了' % name(dic, i))
    for i in range(len(map)):
        if map[i] == 0:
            print('%s 没交' % name(dic, i))
        

def name(dic, info_index):
    name = []
    for k in range(4):
        if 0 < dic['last_choice'][2*k] <= dic['keys_num']:
            name += dic['values'][dic['last_choice'][2*k]-1]
        elif dic['last_choice'][2*k] > dic['keys_num']:
            name += dic['info'][info_index][dic['last_choice'][2*k]-dic['keys_num']-1]
        if k<3:
            name += dic['last_choice'][2*k+1]
    return ''.join(name)
    

start()
