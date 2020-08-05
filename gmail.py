import datetime
import os
import re
import base64
import imaplib
import eml_parser
import pymysql

# my test mail


class gmail(object):
    def __init__(self):
        self.mail_username = 'hahaha@gmail.com'
        self.mail_password = 'password'
        self.conn = pymysql.connect(host='localhost', user='root', password='root', port=3306,
                                    db='bestsellers')  # 有中文要存入数据库的话要加charset='utf8'
        # 创建游标
        self.cursor = self.conn.cursor()

    def receive_email_imap(self):
        server = imaplib.IMAP4_SSL("imap.gmail.com")
        if server is not None:
            res = server.login(self.mail_username, self.mail_password)
            if res != '':
                print('登陆成功', res)
                mail_boxes = []
                boxes=server.list()[1]
                for l in boxes:
                    mail = l.decode('utf-8').rsplit('"/"')[1]
                    mail_pro = mail.strip()
                    mail_boxes.append(mail_pro)
                    # mail_boxes=mail_boxes[1:]
                print('邮箱列表：',mail_boxes)
                if '"INBOX"' in mail_boxes:
                    print('选择邮箱：INBOX')
                    server.select("INBOX", readonly=True)

                    # unseen = server.search(None, 'UNSEEN')[1]  # 选取未读邮件
                    # unseen_list = unseen[0].split()
                    # print('未读邮件数{d}'.format(d=len(unseen_list)))

                    # Recent Seen Answered Flagged Deleted Draft
                    data = server.search(None, 'ALL')[1]#选取所有邮件
                    unseen_list = data[0].split()
                    email_count = len(data[0].split())
                    print('邮件总数 {}'.format(email_count))

                    # email_count
                    for i in range(0, len(unseen_list)):
                        latest_email_uid = unseen_list[i]
                        # print(latest_email_uid)
                        email_data = server.fetch(latest_email_uid, '(RFC822)')[1]
                        raw_email = email_data[0][1]
                        self.parse_emial(raw_email)

    def to_text2(self, subject ,content):
        '''
        :param title: title
        :return:
        '''
        path = self.mk_path()
        path = path + '/' + subject + '.txt'
        print('>>>>已抓取：{content}....<<<<<'.format(content = content[:40]))
        with open(path, 'a', encoding='utf-8') as f:
            f.write(content +'\n'+ '-------'*100)



    def to_text(self, subject ,content):
        '''
        :param title: title
        :return:
        '''
        path = self.mk_path()
        path = path + '/' + subject + '.txt'
        if not os.path.exists(path):
            keys = ['Order ID','Item','Quantity','Return reason','Customer comments','Request received']
            with open(path, 'a', encoding='utf-8') as f:
                f.write('|'.join(keys)+'\n')
        else:
            print('>>>>已保存邮件：{content}....<<<<<'.format(content=content))
            with open(path, 'a', encoding='utf-8') as f:
                f.write('|'.join(content.values()) + '\n')

    def parse_emial(self, msg):
        eml = eml_parser.eml_parser.decode_email_b(msg, include_raw_body=True, include_attachment_data=True)
        header = eml.get('header')
        # print(header)
        #邮件内容
        body = eml.get('body')[0]  # type: list
        subject = header['from']
        content = body['content']
        # print(body)
        if header['from'] == 'order-update@amazon.com':
            info_dict = self.parse_html(content)
            self.to_text(subject, info_dict)
            self.to_mysql(info_dict)
        elif header['from'] == 'do-not-reply@amazon.com' and 'Hello' in content:
            text = re.findall('Hello([\s\S]*)Refund', content)[0].replace('<br>', '')
            text1 = text.replace('\n', '\t')
            text2 = re.sub('<a.*?</a>','', text1)
            self.to_text2(subject, text2)

    def mk_path(self):
        path = os.getcwd()
        work_file = str(datetime.date.today())
        dir_st = os.listdir(path)
        work_path = os.path.join(path, work_file)
        if work_file not in dir_st:
            os.makedirs(work_path)
        return work_path

    def parse_html(self, text):
        info_dict = {}
        try:
            info_dict['Order ID'] = re.findall('Order ID:(.*?)<br>', text)[0]
        except:
            print('此邮件无Order ID')
            info_dict['Order ID'] = 'None'
        try:
            info_dict['Item'] = re.findall('Item:(.*?)<br>', text)[0]
        except:
            print('此邮件无Item')
            info_dict['Item'] = 'None'
        try:
            info_dict['Quantity'] = re.findall('Quantity:(.*?)<br>', text)[0]
        except:
            print('此邮件无Quantity')
            info_dict['Quantity'] = 'None'
        try:
            info_dict['Return reason'] = re.findall('Return reason:(.*?)<br>', text)[0]
        except:
            print('此邮件无Return reason')
            info_dict['Return reason'] = 'None'
        try:
            info_dict['Customer comments'] = re.findall('Customer comments:(.*?)<br>', text)[0]
        except:
            print('此邮件无Customer comments')
            info_dict['Customer comments'] = 'None'
        try:
            info_dict['Request received'] = re.findall('Request received:(.*?)<br>', text)[0]
        except:
            print('此邮件无Request received')
            info_dict['Request received'] = 'None'
        return info_dict

    def to_mysql(self, item):
        # sql语句
        # keys = ','.join(item.keys())
        # values = ','.join(item.values())
        insert_sql = """
        insert into order_update 
        VALUES(%s,%s,%s,%s,%s,%s)
        """
        try:
            # 执行插入数据到数据库操作
            self.conn.ping(reconnect=True)
            self.cursor.execute(insert_sql, (item['Order ID'], item['Item'], item['Quantity'], item['Return reason'],
                                             item['Customer comments'], item['Request received']))
            # 提交，不进行提交无法保存到数据库
            self.conn.commit()
            print('成功保存到数据库')
        except:
            self.conn.ping(reconnect=True)
            print('保存到数据库出错')
            self.cursor.execute(insert_sql, (item['Order ID'], item['Item'], item['Quantity'], item['Return reason'],
                                            item['Customer comments'], item['Request received']))

    #
    # def parse_attachment(self, attachments, save_path):
    #     size = 0
    #     file_list = []
    #
    #     if attachments is None:
    #         return [], size
    #     for file in attachments:
    #         file_name = file.get('filename')
    #         file_size = file.get('size')
    #
    #         size += file_size
    #
    #         # content_header = file.get('content_header')  # type: dict
    #
    #         # save_path=os.path.join(save_path,file_name)
    #         print('存储位置为：', save_path)
    #         file_name = self.save_attachment_file(file.get('raw'), save_path, file.get('filename'))
    #
    #         # 保存所有文件列表
    #         file_list.append(file_name)
    #
    #     return file_list, size

    def save_attachment_file(self, raw, save_path, file_name):

        try:
            raw = base64.b64decode(raw)
            file_path = os.path.join(save_path, file_name)
            with open(file_path, 'wb') as fp:
                fp.write(raw)

        except Exception as e:
            print('解码出错,{}'.format(e))

        return file_path


gmail = gmail()
gmail.receive_email_imap()


