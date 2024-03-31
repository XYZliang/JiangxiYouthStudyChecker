import base64
import hashlib
import json
import logging
import os
import pickle

import pandas as pd
import questionary
import requests
from cryptography.fernet import Fernet
from tqdm import tqdm

logging.basicConfig(level=logging.DEBUG)

# 导出文件夹
output = '导出'
if not os.path.exists('导出'):
    os.mkdir('导出')

# 自定义请求
def doRequest(method, url, headers=None, data=None, cookies=None, max_retries=3, needCookie=False):
    logging.debug(f"Requesting {method} {url}")
    retries = 0
    while retries < max_retries:
        try:
            if method.lower() == 'get':
                response = requests.get(url, headers=headers, params=data, cookies=cookies)
            elif method.lower() == 'post':
                response = requests.post(url, headers=headers, data=data, cookies=cookies)
            else:
                raise ValueError("Unsupported HTTP method")

            # 检查HTTP响应码
            if response.status_code != 200:
                raise Exception(f"HTTP error: {response.status_code}")

            logging.debug(f"Response: {response.text}")
            response_json = response.json()

            # 检查业务逻辑的响应码
            if response_json.get('code') != 200:
                raise Exception(f"API error: {response_json.get('msg')}")
            if needCookie:
                return response_json, response.cookies
            else:
                return response_json  # 如果一切正常，返回响应的JSON数据

        except Exception as e:
            retries += 1
            logging.error(f"Error: {e}. Retrying {retries}/{max_retries}...")

    # 如果重试次数达到上限，抛出异常
    logging.error("Maximum retries reached, operation failed")
    raise Exception("Maximum retries reached, operation failed")


# 登录
def login(account, password):
    logging.info("account:" + account + "正在登录")
    url = "https://hm.jxqingtuan.cn/api-org/user/login"

    payload = json.dumps({
        "account": account,
        "password": password
    })
    headers = {
        'content-type': 'application/json',
    }
    response_json, cookies = doRequest('post', url, headers, payload, needCookie=True)
    logging.debug(response_json)
    if response_json['code'] != 200:
        logging.error("登录失败:" + response_json['msg'])
        exit(response_json['code'])
    token = response_json['data']['token']
    userName = response_json['data']['user']['userName']
    logging.info("登录成功:" + userName)
    logging.debug(cookies)
    # 保存cookies到.ptk文件
    with open('cookies.ptk', 'wb') as file:
        pickle.dump(cookies, file)

    return cookies, token


# 获取未完成列表
def getUserNotFinishRecording(cookies, classid, token, parentId=None, pageSize=1000):
    url = "https://hm.jxqingtuan.cn/api-org/record/getUserNotFinishRecording"
    all_data = []  # 用于存储所有页面的数据

    # 先请求一次获取总页数
    if parentId:
        payload = json.dumps({
            "classId": classid,
            "pageSize": pageSize,
            "currentPage": 1,
            "parentId": parentId
        })
    else:
        payload = json.dumps({
            "classId": classid,
            "pageSize": pageSize,
            "currentPage": 1
        })
    headers = {
        'content-type': 'application/json',
        'Authorization': token,
    }

    response_json = doRequest('post', url, headers, payload, cookies)

    totalPages = response_json['data']['page']['totalPages']

    # 使用tqdm创建进度条
    for page in tqdm(range(1, totalPages + 1), desc="Fetching pages"):
        if parentId:
            payload = json.dumps({
                "classId": classid,
                "pageSize": pageSize,
                "currentPage": page,
                "parentId": parentId
            })
        else:
            payload = json.dumps({
                "classId": classid,
                "pageSize": pageSize,
                "currentPage": page
            })

        response_json = doRequest('post', url, headers, payload, cookies)

        # 添加当前页面的数据
        data = response_json['data']['data']
        all_data.extend(data)

    return all_data


# 获取完成列表
def getUserFinishRecording(cookies, classid, token, parentId=None, pageSize=1000):
    url = "https://hm.jxqingtuan.cn/api-org/record/getUserClassRecord"
    all_data = []  # 用于存储所有页面的数据

    if parentId:
        payload = json.dumps({
            "classId": classid,
            "pageSize": pageSize,
            "currentPage": 1,
            "parentId": parentId
        })
    else:
        payload = json.dumps({
            "classId": classid,
            "pageSize": pageSize,
            "currentPage": 1
        })
    headers = {
        'content-type': 'application/json',
        'Authorization': token,
    }
    response_json = doRequest('post', url, headers, payload, cookies)

    totalPages = response_json['data']['page']['totalPages']

    # 使用tqdm创建进度条
    for page in tqdm(range(1, totalPages + 1), desc="Fetching pages"):
        if parentId:
            payload = json.dumps({
                "classId": classid,
                "pageSize": pageSize,
                "currentPage": page,
                "parentId": parentId
            })
        else:
            payload = json.dumps({
                "classId": classid,
                "pageSize": pageSize,
                "currentPage": page
            })

        response_json = doRequest('post', url, headers, payload, cookies)

        # 添加当前页面的数据
        data = response_json['data']['data']
        all_data.extend(data)

    return all_data


# 获取组织人数信息
def getNumInfo(cookies, token, printInfo=False):
    url = "https://hm.jxqingtuan.cn/api-org/org/getMyOrgNum"

    payload = {}
    headers = {
        'Authorization': token,
        'content-type': 'application/json',
    }

    response_json = doRequest('get', url, headers, payload, cookies)
    if printInfo:
        print("团员数:", response_json['data']['members'])
        print("团干数:", response_json['data']['cadre'])
        print("团支部数:", response_json['data']['orgNum'])
    return response_json


# 获取组织详细信息
def getOrgInfo(cookies, token, printInfo=False):
    url = "https://hm.jxqingtuan.cn/api-org/org/getMyOrgInfo"

    payload = {}
    headers = {
        'Authorization': token,
        'content-type': 'application/json',
    }

    response_json = doRequest('get', url, headers, payload, cookies)
    if printInfo:
        print("上级组织名称:", response_json['data']['parentIdName'])
        print("组织名称:", response_json['data']['orgName'])
        print("组织团员数:", response_json['data']['num'])
    return response_json


# 获取大学习信息
def getClass(cookies, token, printInfo=False):
    url = "https://hm.jxqingtuan.cn/api-org/clazz/getClass"

    payload = {}
    headers = {
        'content-type': 'application/json',
        'Authorization': token,
    }

    response_json = doRequest('get', url, headers, payload, cookies)
    # {"id":68,"title":"2024年第2期","addTime":"2024-03-26 11:00:00","thumb":"","url":"https://h5.cyol.com/special/daxuexi/gq3dj1ys8f/m.html","ext":"","startTime":"2024-03-26 11:00:00","endTime":"2024-04-02 23:59:59","score":20,"pushState":0,"retakes":1,"theme":"青春为中国式现代化挺膺担当","pictureureUrl":"https://oos-cn.ctyunapi.cn/jxgqt/uploadFolder/public/20240326/e1d434a818fb46a295ef0bdec52de634.jpg","retakesPictureUrl":"https://oos-cn.ctyunapi.cn/jxgqt/uploadFolder/public/20240326/5ac0000815ca46628c0ff4f4e9b47c95.png","videoUrl":"","videoCoverImg":"","videoEndImg":"","status":"1","duration":500,"classExamTitle":""}
    if printInfo:
        print("大学习名称:", response_json['data'][0]['title'] + ":" + response_json['data'][0]['theme'])
        print("大学习开始时间:", response_json['data'][0]['startTime'])
        print("大学习结束时间:", response_json['data'][0]['endTime'])
        print("大学习链接:", response_json['data'][0]['url'])
    data_array = response_json['data']
    data_dict = {}
    for item in data_array:
        key = (item['title'] + item['theme']).replace('\n', '')
        value = {
            'id': item['id'],
            'title': item['title'],
            'theme': item['theme'],
            'startTime': item['startTime'],
            'endTime': item['endTime'],
            'url': item['url']
        }
        data_dict[key] = value
    return response_json, data_dict


# 获取组织大学习进度
def getFullSummary(cookies, token, classid, printInfo=False):
    url = "https://hm.jxqingtuan.cn/api-org/record/getOrgLowerClassRecordSummary"

    payload = {
        'classId': classid
    }
    headers = {
        'content-type': 'application/json',
        'Authorization': token,
    }

    response_json = doRequest('get', url, headers, payload, cookies)
    #  {"code":200,"msg":"请求成功","data":{"id":"68","allNum":2398,"num":2097,"title":"2024年第2期","orgName":"软件与物联网工程学院团委","occupancy":87.4500}}
    if printInfo:
        print("大学习标题:", response_json['data']['title'])
        print("大学习总人数:", response_json['data']['allNum'])
        print("大学习已完成人数:", response_json['data']['num'])
        print("大学习完成率:", str(response_json['data']['occupancy']) + "%")

    return response_json


# 获取子组织大学习进度
def getClassSummary(cookies, token, classid):
    url = "https://hm.jxqingtuan.cn/api-org/record/getOrgClassRecord"

    payload = json.dumps({
        "classId": classid
    })
    headers = {
        'content-type': 'application/json',
        'Authorization': token
    }

    response_json = doRequest('post', url, headers, payload, cookies)
    data_array = response_json['data']
    data_dict = {}
    for item in data_array:
        key = item['orgName']
        value = {
            'orgName': item['orgName'],
            'id': item['id'],
            'allNum': item['allNum'],
            'num': item['num'],
            'occupancy': item['occupancy']
        }
        data_dict[key] = value
    return response_json, data_dict


# 获取子组织id
def getClassId(cookies, token, classid):
    url = "https://hm.jxqingtuan.cn/api-org/record/getOrgClassRecord"

    payload = json.dumps({
        "classId": classid
    })
    headers = {
        'content-type': 'application/json',
        'Authorization': token
    }

    response_json = doRequest('post', url, headers, payload, cookies)
    data_array = response_json['data']
    data1_dict = {}
    data2_dict = {}
    for item in data_array:
        id = item['id']
        name = item['orgName']
        data1_dict[id] = name
        data2_dict[name] = {
            'orgName': item['orgName'],
            'id': item['id'],
            'allNum': item['allNum'],
            'num': item['num'],
            'occupancy': item['occupancy']
        }
    return data1_dict, data2_dict


# 加密和解密的函数
def encrypt_message(message, key):
    f = Fernet(key)
    encrypted_message = f.encrypt(message.encode())
    return encrypted_message


def decrypt_message(encrypted_message, key):
    f = Fernet(key)
    decrypted_message = f.decrypt(encrypted_message).decode()
    return decrypted_message


# 保存和加载凭据的函数
def save_credentials(filename, account, password, key):
    encrypted_account = encrypt_message(account, key)
    encrypted_password = encrypt_message(password, key)
    with open(filename, 'wb') as f:
        f.write(encrypted_account + b'\n' + encrypted_password)


def load_credentials(filename, key):
    with open(filename, 'rb') as f:
        encrypted_account, encrypted_password = [line.strip() for line in f.readlines()]
    account = decrypt_message(encrypted_account, key)
    password = decrypt_message(encrypted_password, key)
    return account, password


# 嵌套菜单的函数
def studyMenu(cookies, token):
    classInfoJson, classInfoDict = getClass(cookies, token, printInfo=False)
    classList = list(classInfoDict.keys())
    classList[0] = classList[0] + "（最近一期/当前大学习）"
    classChoice = questionary.select("请选择一个需要查询的期数:", choices=classList).ask()
    classIndex = classList.index(classChoice)
    classKey = list(classInfoDict.keys())[classIndex]
    classId = classInfoDict[classKey]['id']
    print("大学习名称:", classInfoJson['data'][classIndex]['title'] + ":" + classInfoJson['data'][classIndex]['theme'])
    print("大学习开始时间:", classInfoJson['data'][classIndex]['startTime'])
    print("大学习结束时间:", classInfoJson['data'][classIndex]['endTime'])
    print("大学习链接:", classInfoJson['data'][classIndex]['url'])
    # 查询完成率
    getFullSummary(cookies, token, classId, printInfo=True)
    choices = [
        '导出全部完成情况名单',
        '导出全部未完成情况名单',
        '查看并导出特定子团支部完成情况名单',
        '查看并导出特定子团支部未完成情况名单',
        '查看并导出各个子团支部完成情况统计',
        '返回主菜单'
    ]
    while True:
        choice = questionary.select("请选择一个需要查询的期数:", choices=choices).ask()
        fileTitle = None
        # 准备数据
        if choice == '导出全部完成情况名单' or choice == '导出全部未完成情况名单':
            id2name_dict, _ = getClassId(cookies, token, classId)
            if choice == '导出全部完成情况名单':
                data = getUserFinishRecording(cookies, classId, token)
                fileTitle = '完成名单'
                df = pd.DataFrame(data)
                df['组织'] = df[['lev1', 'lev2', 'lev3', 'lev4']].apply('-'.join, axis=1)
                df.rename(columns={'addTime': '学习时间', 'username': '姓名'}, inplace=True)
                # 定义要隐藏的列
                columns_to_hide = ['id', 'classId', 'score', 'lev1', 'lev2', 'lev3', 'lev4', 'userid', 'nid', 'subOrg',
                                   'nid1', 'nid2', 'nid3', 'status', 'studyTime']

                # 导出DataFrame到Excel
                filename = f'./导出/{fileTitle}-{classInfoJson["data"][classIndex]["title"]}-{pd.Timestamp.now().strftime("%Y%m%d%H%M%S")}.xlsx'
                with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')

                    # 获取工作表
                    worksheet = writer.sheets['Sheet1']

                    # 遍历需要隐藏的列
                    for col_name in columns_to_hide:
                        # 获取列的索引
                        col_idx = pd.Index(df.columns).get_loc(col_name)
                        # 隐藏列，不设置宽度和格式
                        worksheet.set_column(col_idx, col_idx, None, options={'hidden': True})
            else:
                data = getUserNotFinishRecording(cookies, classId, token)
                fileTitle = '未完成名单'
                # 加载数据到DataFrame
                df = pd.DataFrame(data)

                # 生成团支部列
                def map_area(row):
                    for i in range(1, 6):
                        try:
                            area_id = row[f'areaid{i}']
                            if area_id in id2name_dict:
                                return id2name_dict[area_id]
                        except KeyError:
                            continue
                    return ""

                df['团支部'] = df.apply(map_area, axis=1)

                # 重命名username列为姓名
                df.rename(columns={'username': '姓名'}, inplace=True)

                # 创建一个Excel写入器
                # 名称为 fileTitle-classInfoJson['data'][classIndex]['title']-当前时间.xlsx
                with pd.ExcelWriter(
                        f'./导出/{fileTitle}-{classInfoJson["data"][classIndex]["title"]}-{pd.Timestamp.now().strftime("%Y%m%d%H%M%S")}.xlsx') as writer:
                    df.to_excel(writer, index=False)

                    # 获取工作表
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']

                    # 设置除了团支部和姓名的所有列为隐藏
                    for idx, col in enumerate(df.columns):
                        if col not in ['团支部', '姓名']:
                            worksheet.set_column(idx, idx, None, None, {'hidden': True})
        elif choice == '查看并导出特定子团支部完成情况名单' or choice == '查看并导出特定子团支部未完成情况名单':
            # 选择团支部
            id2name_dict, class_dict = getClassId(cookies, token, classId)
            listChoices = list()
            for key, value in class_dict.items():
                content = value['orgName'] + "（" + str(value['num']) + "/" + str(value['allNum']) + " " + str(
                    value['occupancy']) + "%）"
                listChoices.append(content)
            listChoices.append("返回主菜单")
            classChoice = questionary.select("请选择一个需要查询的团支部:", choices=listChoices).ask()
            if classChoice == "返回主菜单":
                return
            class_Index = listChoices.index(classChoice)
            class_Key = list(class_dict.keys())[class_Index]
            class_Id = class_dict[class_Key]['id']
            if choice == '查看并导出特定子团支部完成情况名单':
                data = getUserFinishRecording(cookies, classId, token, parentId=class_Id)
                fileTitle = class_Key + '完成名单'
                df = pd.DataFrame(data)
                df['组织'] = df[['lev1', 'lev2', 'lev3', 'lev4']].apply('-'.join, axis=1)
                df.rename(columns={'addTime': '学习时间', 'username': '姓名'}, inplace=True)
                # 定义要隐藏的列
                columns_to_hide = ['id', 'classId', 'score', 'lev1', 'lev2', 'lev3', 'lev4', 'userid', 'nid', 'subOrg',
                                   'nid1', 'nid2', 'nid3', 'status', 'studyTime']

                # 导出DataFrame到Excel
                filename = f'./导出/{fileTitle}-{classInfoJson["data"][classIndex]["title"]}-{pd.Timestamp.now().strftime("%Y%m%d%H%M%S")}.xlsx'
                with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')

                    # 获取工作表
                    worksheet = writer.sheets['Sheet1']

                    # 遍历需要隐藏的列
                    for col_name in columns_to_hide:
                        # 获取列的索引
                        col_idx = pd.Index(df.columns).get_loc(col_name)
                        # 隐藏列，不设置宽度和格式
                        worksheet.set_column(col_idx, col_idx, None, options={'hidden': True})
            else:
                data = getUserNotFinishRecording(cookies, classId, token, parentId=class_Id)
                fileTitle = class_Key + '未完成名单'
                # 加载数据到DataFrame
                df = pd.DataFrame(data)

                # 重命名username列为姓名
                df.rename(columns={'username': '姓名'}, inplace=True)

                # 创建一个Excel写入器
                # 名称为 fileTitle-classInfoJson['data'][classIndex]['title']-当前时间.xlsx
                with pd.ExcelWriter(
                        f'./导出/{fileTitle}-{classInfoJson["data"][classIndex]["title"]}-{pd.Timestamp.now().strftime("%Y%m%d%H%M%S")}.xlsx') as writer:
                    df.to_excel(writer, index=False)

                    # 获取工作表
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']

                    # 设置除了姓名的所有列为隐藏
                    for idx, col in enumerate(df.columns):
                        if col not in ['姓名']:
                            worksheet.set_column(idx, idx, None, None, {'hidden': True})
        elif choice == '查看并导出各个子团支部完成情况统计':
            fileTitle = "各团支部完成情况统计"
            data, _ = getClassSummary(cookies, token, classId)
            df = pd.DataFrame(data['data'])
            # 重命名列
            df = df.rename(columns={
                'allNum': '总人数',
                'num': '已学习人数',
                'orgName': '团支部',
                'occupancy': '学习率'
            })
            # 重新排序列
            df = df[['id', '团支部', '已学习人数', '总人数', '学习率']]
            # 根据团支部排序
            df = df.sort_values(by='团支部')
            # 导出DataFrame到Excel
            with pd.ExcelWriter(f'./导出/{fileTitle}-{classInfoJson["data"][classIndex]["title"]}-{pd.Timestamp.now().strftime("%Y%m%d%H%M%S")}.xlsx', engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                # 获取工作表
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                # 隐藏id列
                worksheet.set_column('A:A', None, None, {'hidden': True})
        elif choice == '返回主菜单':
            return
        # 是否返回主菜单
        returnMain = questionary.confirm("导出成功！是否返回主菜单？").ask()
        if returnMain:
            return


def main_menu():
    isLogin = False
    credentials_file = 'account.data'
    # 你的原始密钥
    password = 'JiangxiYouthStudyMaker'
    # 使用SHA-256哈希函数来生成一个固定长度的字节序列
    hash = hashlib.sha256(password.encode()).digest()
    # 将生成的哈希值转换为Base64编码以获得Fernet密钥
    key = base64.urlsafe_b64encode(hash)

    while True:
        if not isLogin and os.path.exists(credentials_file):
            use_saved = questionary.confirm("发现保存的账号信息，是否直接登录？").ask()
            if use_saved:
                account, password = load_credentials(credentials_file, key)
                cookies, token = login(account, password)
                isLogin = True
            else:
                os.remove(credentials_file)  # 删除旧的凭据文件
                # 重置
                continue

        if not isLogin:
            account = questionary.text("请输入账号：").ask()
            password = questionary.password("请输入密码：").ask()
            cookies, token = login(account, password)
            isLogin = True

            save = questionary.confirm("是否保存密码？").ask()
            if save:
                save_credentials(credentials_file, account, password, key)

        choices = [
            '获取组织人数信息',
            '获取组织详细信息',
            '查看/导出大学习信息',
            '完成大学习',
            '退出'
        ]
        choice = questionary.select("请选择一个操作:", choices=choices).ask()

        if choice == '退出':
            break
        elif choice == '获取组织人数信息':
            getNumInfo(cookies, token, printInfo=True)
        elif choice == '获取组织详细信息':
            getOrgInfo(cookies, token, printInfo=True)
        elif choice == '查看/导出大学习信息':
            studyMenu(cookies, token)
        elif choice == '完成大学习':
            exit()


if __name__ == "__main__":
    main_menu()
