from wcferry import Wcf,WxMsg
from queue import Empty
from threading import Thread
import re
import pandas as pd

#[{'wxid': '45398101613@chatroom', 'code': '', 'remark': '', 'name': '智能技术班', 'country': '', 'province': '', 'city': '', 'gender': ''}]
#{'wxid': '43661614402@chatroom', 'code': '', 'remark': '', 'name': '测试机器人', 'country': '', 'province': '', 'city': '', 'gender': ''}
new_score=0
row=None

cs_zrid='wxid_gde1pf35qb7i19'
zrid='wxid_wzoy9mr364a322'
bqid='45398101613@chatroom'
cs_bqid='43661614402@chatroom'
wcf=Wcf()#hook微信客户端
wc_log=wcf.is_login()#登录检测
df = pd.read_excel('智能管理班操行分.xlsx')

#获取被@同学的wxid
def find_classmate(msg):
    match = re.search(r'\[wxid_(.*?)]', str(msg))
    classmate='wxid_'+match.group(1)
    return classmate

#操行分查询
df = pd.read_excel('智能管理班操行分.xlsx')

def find_all(msg):
    result_df = df[['Name', '操行分']]
    wcf.send_text(str(result_df.to_string(index=False)),bqid)#0000000000
    print('查询成功')
    
#重新开始记录
def start():
    bq_member=wcf.get_chatroom_members(cs_bqid)
    df = pd.DataFrame.from_dict(bq_member, orient='index', columns=['Name'])
    df.to_excel('智能管理班操行分.xlsx', index_label='wxid')
    df['操行分'] = 0
    df.to_excel('智能管理班操行分.xlsx', index_label='wxid')

#识别列表中的群

wcf_rooms = []#存储群聊消息
def bq_get():
    for contact in wcf.get_contacts():
        if contact['name']=='智能技术班':
            wcf_rooms.append(contact)


def u_info():#打印登录账号消息
    user_info=wcf.get_user_info()
    print('wxid：{}'.format(user_info['wxid']))
    print('微信昵称：{}'.format(user_info['name']))

def processMsg(msg: WxMsg):#接收打印消息
        print(msg.content)

#操行分加减
def update_score(msg):
    # 根据 wxid 找到对应的行
    row = df[df['wxid'] == find_classmate(msg)]
    if row.shape[0] > 0:
        # 获取当前操行分
        current_score = int(row['操行分'].values[0])
        # 加减分

        match = re.search(r'加(\d+)分', msg.content)
        if match:
            x = int(match.group(1))
            new_score = current_score + x
            print('成功加{}分'.format(x))
        else:
            match = re.search(r'减(\d+)分', msg.content)
            x = int(match.group(1))
            new_score = current_score - x
            print('成功减{}分'.format(x))

        # 发送加分信息
        wcf.send_text(f'{row["Name"].values[0]}, 当前操行分：{new_score}', bqid)

        return new_score, row
    else:
        print(f'未找到微信 ID 为 {find_classmate(msg)} 的记录')

#接收消息
def enableReceivingMsg():
    def innerWcFerryProcessMsg():
        while wcf.is_receiving_msg():#is_receiving_msg方法用于检查是否正在接收消息。它可能是一个布尔值方法，返回True表示正在接收消息，返回False表示当前没有接收消息。
            try:
                msg = wcf.get_msg()#获取消息
                processMsg(msg)
                if msg.from_group:
                    # if True:
                    if msg.sender==zrid:
                        if msg.content == '操行分查询':
                            find_all(msg)
                            
                        if msg.content=="重启程序":
                            bq_get()
                            start()
                            ret=wcf.send_text("重启成功",msg.sender)#0000000000000000000
                        n_r=update_score(msg)
                        # 更新操行分
                        df.loc[n_r[1].index[0], '操行分'] = n_r[0]
                        df.to_excel('智能管理班操行分.xlsx')
                        find_classmate(msg)
            except Empty:
                continue
            except Exception as e:
                print(f"ERROR: {e}")
    wcf.enable_receiving_msg()#会启动接收消息的机制，使得程序能够开始接收和处理消息
    Thread(target=innerWcFerryProcessMsg, name="ListenMessageThread", daemon=True).start()


u_info()
enableReceivingMsg()
wcf.keep_running()

