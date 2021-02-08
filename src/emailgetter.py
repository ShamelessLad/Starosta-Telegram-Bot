import imaplib #pip install python-imap
import email #уже есть в стандартном наборе
import os, string, random, re, base64, quopri

class EmailLoginError(Exception):
    def __init__(self):
        self.message = "Incorrect email address or password"
    def __str__(self):
        return "LoginError: {0}".format(self.message)

class EmailGetter:
    email_address = ""
    password = ""
    imap_host = "imap.gmail.com"
    imap_port = 993
    
    def __init__(self, **kwargs):
        for key, value in kwargs.items():
            if key == "email_address":
                self.email_address = value
            if key == "password":
                self.password = value
            if key == "imap_host":
                self.imap_host = value
            if key == "imap_port":
                self.imap_port = value
    
    def get_last_messages(self, last_messages_num: int) -> list:
        """
        Возвращает список, состоящий из последних last_messages_num сообщений
        Каждое сообщение характеризуется словарем формата:
        {
            'from': str(<имя-отправителя>),
            'subject': str(<тема-письма>),
            'plain_text': str(<Текст сообщения>),
            'attachments_dirs': [список, путей, на, вложения]
            'id': id сообщения
        }
        """
        return_list = []

        try:
            mail = imaplib.IMAP4_SSL(self.imap_host, self.imap_port)
            mail.login(self.email_address, self.password)
            status, messages = mail.select("INBOX")
            if status != "OK":
                print("smth wrong with email")
                return
            messages = int(messages[0])
        except (imaplib.IMAP4_SSL.error, imaplib.IMAP4.error) as e: #я просто хз какой тут класс у исключения
            return #какой то дебаггерский вывод, что залогиниться не удалось


        for i in range(messages, messages - last_messages_num, -1):
            res, msg = mail.fetch(str(i), "(RFC822)")
            for response in msg:
                if isinstance(response, tuple):
                    dict_message = {
                        "from": "",
                        "subject": "",
                        "plain_text": "",
                        "attachments_dirs": [],
                        "id": int(i) #на всякий и тут получаем id, лишним не будет
                    }
                    msg = email.message_from_bytes(response[1])

                    subject, encoding = email.header.decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding)

                    from_who, encoding = email.header.decode_header(msg.get("From"))[0]
                    if isinstance(from_who, bytes):
                        from_who = from_who.decode(encoding)
                    
                    dict_message["from"] = from_who
                    dict_message["subject"] = subject

                    if msg.is_multipart():
                        for part in msg.walk():
                            content_type        = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))

                            try:
                                body = part.get_payload(decode=True).decode()
                            except:
                                pass #вообще хз какие тут эксепшены, кек

                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                dict_message["plain_text"] = body
                            elif "attachment" in content_disposition:
                                filename = part.get_filename()

                                #манипуляции с регулярками, потому что в библиотеке не предусмотрели эту расшифровку
                                encoded_filename_regex = r'=\?{1}(.+)\?{1}([B|Q])\?{1}(.+)\?{1}=' #регулярка для закодированного текста
                                charset, encoding, encoded_text = re.match(encoded_filename_regex, filename).groups()
                                if encoding == 'B': #в бинарном виде в base64
                                    bytes_filename = base64.b64decode(encoded_text)
                                if encoding == 'Q': #в бинарном виде в MIME
                                    bytes_filename = quopri.decodestring(encoded_text)
                                #преобразовываем байты в указанную кодировку
                                #обычно это всегда utf-8, но стоит указать что там реально
                                filename = bytes_filename.decode(charset)
                                
                                if filename:
                                    """ Под каждый файл создается отдельная папка, потому что:
                                        1) Чтобы не нарушались названия файлов
                                        2) В разных сообщениях могут быть вложения с одинаковыми названиями,
                                        поэтому нельзя делать под один проход по почте одну папку для вложений
                                    """
                                    if not os.path.isdir("attachments"):
                                        os.mkdir("attachments")
                                    random_foldername = os.path.join("attachments", "".join([random.choice(string.ascii_letters + string.digits) for n in range(10)]))
                                    os.mkdir(random_foldername)
                                    filepath = os.path.join(random_foldername, filename)

                                    with open(filepath, "wb") as attachment_file:
                                        attachment_file.write(part.get_payload(decode=True))
                                    attachment_file.close()
                                    dict_message["attachments_dirs"].append(filepath)
                                    #подчищать вложения на совести того, кто вызвал функцию, потому что не знаем, когда нужно удалять
                    else:
                        pass #здесь получать html форму (если в будущем понадобится)
                    return_list.append(dict_message)

        mail.close()
        mail.logout()
        return return_list
    
    def get_newer_messages(self, latest_message_id: int) -> list:
        """
        Возвращает список сообщений, которые новее чем сообщение с id latest_message_id
        Каждое сообщение характеризуется словарем формата:
        {
            'from': str(<имя-отправителя>),
            'subject': str(<тема-письма>),
            'plain_text': str(<Текст сообщения>),
            'attachments_dirs': [список, путей, на, вложения]
            'id': id сообщения
        }
        """
        return_list = []

        try:
            mail = imaplib.IMAP4_SSL(self.imap_host, self.imap_port)
            mail.login(self.email_address, self.password)
            status, messages = mail.select("INBOX")
            if status != "OK":
                print("smth wrong with email")
                return
            messages = int(messages[0])
        except (imaplib.IMAP4_SSL.error, imaplib.IMAP4.error) as e: #я просто хз какой тут класс у исключения
            return #какой то дебаггерский вывод, что залогиниться не удалось


        for i in range(messages, latest_message_id, -1):
            res, msg = mail.fetch(str(i), "(RFC822)")
            for response in msg:
                if isinstance(response, tuple):
                    dict_message = {
                        "from": "",
                        "subject": "",
                        "plain_text": "",
                        "attachments_dirs": [],
                        "id": int(i)
                    }
                    msg = email.message_from_bytes(response[1])

                    subject, encoding = email.header.decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding)

                    from_who, encoding = email.header.decode_header(msg.get("From"))[0]
                    if isinstance(from_who, bytes):
                        from_who = from_who.decode(encoding)
                    
                    dict_message["from"] = from_who
                    dict_message["subject"] = subject

                    if msg.is_multipart():
                        for part in msg.walk():
                            content_type        = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))

                            try:
                                body = part.get_payload(decode=True).decode()
                            except:
                                pass #вообще хз какие тут эксепшены, кек

                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                dict_message["plain_text"] = body
                            elif "attachment" in content_disposition:
                                filename = part.get_filename()

                                try:
                                    #манипуляции с регулярками, потому что в библиотеке не предусмотрели эту расшифровку
                                    encoded_filename_regex = r'=\?{1}(.+)\?{1}([B|Q])\?{1}(.+)\?{1}=' #регулярка для закодированного текста
                                    charset, encoding, encoded_text = re.match(encoded_filename_regex, filename).groups()
                                    if encoding == 'B': #в бинарном виде в base64
                                        bytes_filename = base64.b64decode(encoded_text)
                                    if encoding == 'Q': #в бинарном виде в MIME
                                        bytes_filename = quopri.decodestring(encoded_text)
                                    #преобразовываем байты в указанную кодировку
                                    #обычно это всегда utf-8, но стоит указать что там реально
                                    filename = bytes_filename.decode(charset)
                                except AttributeError:
                                    pass
                                
                                if filename:
                                    """ Под каждый файл создается отдельная папка, потому что:
                                        1) Чтобы не нарушались названия файлов
                                        2) В разных сообщениях могут быть вложения с одинаковыми названиями,
                                        поэтому нельзя делать под один проход по почте одну папку для вложений
                                    """
                                    if not os.path.isdir("attachments"):
                                        os.mkdir("attachments")
                                    random_foldername = os.path.join("attachments", "".join([random.choice(string.ascii_letters + string.digits) for n in range(10)]))
                                    os.mkdir(random_foldername)
                                    filepath = os.path.join(random_foldername, filename)

                                    with open(filepath, "wb") as attachment_file:
                                        attachment_file.write(part.get_payload(decode=True))
                                    attachment_file.close()
                                    dict_message["attachments_dirs"].append(filepath)
                                    #подчищать вложения на совести того, кто вызвал функцию, потому что не знаем, когда нужно удалять
                    else:
                        pass #здесь получать html форму (если в будущем понадобится)
                    return_list.append(dict_message)

        mail.close()
        mail.logout()
        return return_list

    def get_last_message_id(self):
        try:
            mail = imaplib.IMAP4_SSL(self.imap_host, self.imap_port)
            mail.login(self.email_address, self.password)
            status, messages = mail.select("INBOX")
            if status != "OK":
                print("smth wrong with email")
                return
        except (imaplib.IMAP4_SSL.error, imaplib.IMAP4.error) as e: #я просто хз какой тут класс у исключения
            print(e)
            return #какой то дебаггерский вывод, что залогиниться не удалось
        messages = int(messages[0])
        mail.close()
        mail.logout()
        return messages

def main(): #для тестирования
    eg = EmailGetter(
        email_address = "sampleaddress@domain.com",
        password      = "verycomplexpassword"
    )
    print(eg.get_newer_messages(1337))

if __name__ == "__main__":
    main()
