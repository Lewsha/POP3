import socket
import ssl
import argparse
import base64
import re
import os
import shutil


TIMEOUT = 8  # Наш таймаут
SSL_POP3_PORT = 995  # номер защищенного порта pop3


class POP3Exception(Exception):
    """Наша собственная ошибка, ееее!"""
    def __init__(self, message):
        super().__init__(message)  # Тупо super'имся от родителя с нужным сообщением


class MailStruct:
    """Класса объета письма (данные о письме, не само письмо)"""
    def __init__(self, mail_from, mail_to, mail_subject,
                 mail_date, mail_size, original_headers):
        self.mail_from = mail_from  # от кого письмо
        self.mail_to = mail_to  # кому письмо
        self.mail_subject = mail_subject  # тема письма
        self.mail_date = mail_date  # дата отправки
        self.mail_size = mail_size  # размер письма
        self.original_headers = original_headers

    def __repr__(self):
        """Переопределяем представление объекта"""
        headers = ('From: {}\n'
                   'To: {}\n'
                   'Subject: {}\n'
                   'Date: {}\n'
                   'Mail Size: {} bytes\n').format(
            self.mail_from,
            self.mail_to,
            self.mail_subject,
            self.mail_date,
            self.mail_size)  # вставляем нужные данные в нужные места
        return headers  # возвращаем представление


def print_help():
    """чисто выводим список команд"""
    print(
        """#######
        Список команд:
EXIT - выйти из программы
HELP - вызвать список команд
LIST - получить список писем с основной информацией 
        (от кого, кому, когда пришло, тема письма)
TOP [номер письма] [количество строк письма] - получить заголовки указанного письма
        (выведется лишь основная информация) и указанное число строк письма.
RECV [номер письма] - скачать указанное письмо
RECV ALL - скачать все письма
#######

        """
    )


def print_list(letters_info):
    """Просто функция вывода инфы по письмам"""
    for letter_number in range(len(letters_info)):
        print('\t\t\t Message №  {}'.format(letter_number + 1))
        try:
            print(letters_info[letter_number], '\n')
        except UnicodeEncodeError:
            print('Can\'t print this')  # если что-то не вышло с декодировкой, то принтуем инфу об этом


def print_top(channel, letter_number, lines_count):
    """Функция, позволяющая вывести несколько первых строк письм
    Выводит именно строки текста письма, куски файлов отбрасывает"""
    """Если пользователь не передал количество строк, отправляем команду с одним аргументом"""
    if lines_count is None:
        send(channel, 'TOP {}'.format(letter_number))  # посылаем команду TOP с номером нужного письма
    else:
        """А в обычном случае - с двумя"""
        send(channel, 'TOP {} {}'.format(letter_number, lines_count))  # посылаем команду TOP с номером нужного письма
    letter = '\n'.join(recv_multiline(channel))  # возвращаем полученное сообщение
    headers, *body = letter.split('\n\n')  # парсим на тело и заголовки
    body = '\n\n'.join(body)  # соединяем тело в одну строку
    file_name_req_exp = re.compile('filename="(.+?)"', flags=re.IGNORECASE)  # регулярка для имени файла
    encoding_req_exp = re.compile('Content-Transfer-Encoding: (.+)', flags=re.IGNORECASE)  # регулярка для кодировки
    base64_req_exp = re.compile(r'=\?(.*?)\?b', flags=re.IGNORECASE)  # регулярка на случай base64
    content_type_req_exp = re.compile('Content-Type: (.+?)/(.+?)[;,\n]', flags=re.IGNORECASE)  # регулярка для MIME-типа
    boundary_req_exp = re.compile('boundary="(.*?)"')  # регулярка для разделителя

    boundary = re.search(boundary_req_exp, headers)  # ищем разделитель
    if not boundary:
        # если не нашли, искусственно создаем
        boundary = 'boundary'
    else:
        boundary = boundary.group(1)

    """словарь MIME-типов, чтобы можно было понять, какое расширение у файла"""
    extensions = {'image': {'png': 'png', 'gif': 'gif', 'x-icon': 'ico', 'jpeg': ['jpeg', 'jpg'],
                            'svg+xml': 'svg', 'tiff': 'tiff', 'webp': 'webp', 'bmp': 'bmp'},
                  'video': {'x-msvideo': 'avi', 'mpeg': ['mp3', 'mpeg'], 'ogg': 'ogv', 'webm': 'webm'},
                  'application': {'zip': 'zip', 'xml': 'xml', 'octet-stream': 'bin', 'x-bzip': 'bz',
                                  'msword': 'doc', 'epub+zip': 'epub', 'javascript': 'js', 'json': 'json',
                                  'pdf': 'pdf', 'vnd.ms-powerpoint': 'ppt', 'x-rar-compressed': 'rar',
                                  'x-sh': 'sh', 'x-tar': 'tar'},
                  'audio': {'x-wav': 'wav', 'ogg': 'oga'},
                  'text': {'css': 'css', 'csv': 'csv', 'html': 'html', 'plain': 'txt'}}

    if 'Content-Type: multipart/mixed' in headers:
        body = body.split('--' + boundary + '\n')
        body[-1] = body[-1][:len(body[-1]) - len(boundary) - 4]
        if body[0] == '\n':
            body.pop(0)

        for attachment in body:
            attachment_data = attachment.split('\n\n')
            if len(attachment_data) < 2:
                continue

            attachment_data = attachment_data[1]

            encoding = encoding_req_exp.search(attachment).group(1)

            file_name = file_name_req_exp.search(attachment)

            if not file_name:
                content_info = content_type_req_exp.search(attachment)
                c_type = content_info.group(1)
                extension = extensions[c_type][content_info.group(2)]
                file_name = "Letter № {}.{}".format(letter_number, extension)

            else:
                continue

            if encoding == "base64":
                print(base64.b64decode(attachment_data))
            else:
                print(attachment_data)
    else:
        content_info = content_type_req_exp.search(headers)
        c_type = content_info.group(1)
        extension = extensions[c_type][content_info.group(2)]
        encoding = encoding_req_exp.search(headers).group(1)

        if encoding == "base64":
            print(base64.b64decode(body))
        else:
            print(body)


def recv_letter(channel, letter_number, letter_info):
    """Функция, скачивающая письмо с сервера"""
    send(channel, 'RETR {}'.format(letter_number))  # посылаем команду RETR с номером нужного письма
    letter = '\n'.join(recv_multiline(channel))  # возвращаем полученное сообщение
    headers, *body = letter.split('\n\n')
    body = '\n\n'.join(body)

    extensions = {'image': {'png': 'png', 'gif': 'gif', 'x-icon': 'ico', 'jpeg': ['jpeg', 'jpg'],
                            'svg+xml': 'svg', 'tiff': 'tiff', 'webp': 'webp', 'bmp': 'bmp'},
                  'video': {'x-msvideo': 'avi', 'mpeg': ['mp3', 'mpeg'], 'ogg': 'ogv', 'webm': 'webm'},
                  'application': {'zip': 'zip', 'xml': 'xml', 'octet-stream': 'bin', 'x-bzip': 'bz',
                                  'msword': 'doc', 'epub+zip': 'epub', 'javascript': 'js', 'json': 'json',
                                  'pdf': 'pdf', 'vnd.ms-powerpoint': 'ppt', 'x-rar-compressed': 'rar',
                                  'x-sh': 'sh', 'x-tar': 'tar'},
                  'audio': {'x-wav': 'wav', 'ogg': 'oga'},
                  'text': {'css': 'css', 'csv': 'csv', 'html': 'html', 'plain': 'txt'}}

    file_name_req_exp = re.compile('filename="(.+?)"', flags=re.IGNORECASE)
    encoding_req_exp = re.compile('Content-Transfer-Encoding: (.+)', flags=re.IGNORECASE)
    base64_req_exp = re.compile(r'=\?(.*?)\?b', flags=re.IGNORECASE)
    content_type_req_exp = re.compile('Content-Type: (.+?)/(.+?)[;,\n]', flags=re.IGNORECASE)
    boundary_req_exp = re.compile('boundary="(.*?)"')

    boundary = re.search(boundary_req_exp, letter_info.original_headers)
    if not boundary:
        boundary = 'boundary'
    else:
        boundary = boundary.group(1)

    folder = 'Letter № {}'.format(letter_number)
    try:
        if os.path.exists(folder):
            shutil.rmtree(folder)
        os.mkdir(folder)
    except Exception as er:
        print(er)

    try:
        with open('{}/Letter № {} info.{}.txt'.format(folder, letter_number, boundary), mode='w') as file:
            file.write(repr(letter_info))
    except Exception as er:
        print(er)

    if 'Content-Type: multipart/mixed' in letter_info.original_headers:

        body = body.split('--' + boundary + '\n')
        body[-1] = body[-1][:len(body[-1]) - len(boundary) - 4]
        if body[0] == '\n':
            body.pop(0)

        for attachment in body:
            attachment_data = attachment.split('\n\n')
            if len(attachment_data) < 2:
                continue

            attachment_data = attachment_data[1]

            encoding = encoding_req_exp.search(attachment).group(1)

            file_name = file_name_req_exp.search(attachment)

            if not file_name:
                content_info = content_type_req_exp.search(attachment)
                c_type = content_info.group(1)
                extension = extensions[c_type][content_info.group(2)]
                file_name = "Letter № {}.{}".format(letter_number, extension)

            else:
                file_name = file_name.group(1)

            if base64_req_exp.match(file_name):
                file_name = decode_inline_base64(file_name)

            with open('{}/{}'.format(folder, file_name), mode='wb') as file:
                if encoding == "base64":
                    file.write(base64.b64decode(attachment_data))
                else:
                    file.write(attachment_data.encode())
    else:
        content_info = content_type_req_exp.search(letter_info.original_headers)
        c_type = content_info.group(1)
        extension = extensions[c_type][content_info.group(2)]
        file_name = "Letter № {}.{}".format(letter_number, extension)
        encoding = encoding_req_exp.search(letter_info.original_headers).group(1)

        with open('{}/{}'.format(folder, file_name), mode='wb') as file:
            if encoding == "base64":
                file.write(base64.b64decode(body))
            else:
                file.write(body.encode())


def recv_all(channel, letters_info):
    for letter_number in range(1, len(letters_info) + 1):
        print(letter_number)
        recv_letter(channel, letter_number, letters_info[letter_number - 1])
    print('Successful')


def get_info(channel):
    letters_structs = []  # создаем лист под инфу о письмах
    send(channel, 'LIST')  # посылаем команду LIST для получения списка писем
    """Из полученного ответа вынимаем интересующую информацию"""
    for line in recv_multiline(channel):
        msg_id, msg_size = map(int, line.split())  # из строки забираем номер письма и его размер
        send(channel, 'TOP {} 0'.format(msg_id))
        headers = recv_multiline(channel)
        headers = '\n'.join(headers)
        important_headers = ['From', 'To', 'Subject', 'Date']
        important_headers_values = []
        for header in important_headers:
            important_headers_values.append(
                find_header(header, headers) or 'unknown')
        important_headers_values = list(
            map(decode_inline_base64, important_headers_values))
        msg_from, msg_to, msg_subj, msg_date \
            = important_headers_values  # парсим сообщение и забираем нужные данные о письме
        letters_structs.append(MailStruct(
            msg_from, msg_to, msg_subj, msg_date, msg_size, headers))  # создаем структуру из полученных данных
    return letters_structs  # возвращаем полученные данные


def regexp_post_processing(s):
    """Функция для, улучшения читаемости получаемых заголовков письма
    ПОЯСНЯЮ! Заголовок может быть сделан в несколько строк, с ненужными пробелами и т.д."""
    flags = re.IGNORECASE | re.MULTILINE | re.DOTALL  # выставляем флаги для регулярного выражения
    return re.sub(r'^\s*', '', s.strip(), flags=flags).replace('\n', '')  # возвращаем, по сути, то, что получили


def find_header(header, headers):
    """Функция для вытяжки из письма нужных нам заголовков"""
    flags = re.IGNORECASE | re.MULTILINE | re.DOTALL  # устанавливаем флаги для регулярного выражения
    match = re.search(r'^{}:(.*?)(?:^\w|\Z|\n\n)'
                      .format(header), headers, flags)  # ищем заголовок в письме
    if match:
        return regexp_post_processing(match.group(1))  # если нашли, возвращаем грамотно скомпанованные заголовки


def decode_inline_base64(value):
    """Функция декодирования base64"""
    """Дополнительная функция, создающая подстроку, на которую мы заменим подстроку в base64"""
    def replace(match):
        encoding = match.group(1).lower()  # в первой группе у нас будет содержаться кодировка (напр., =?utf-8?b?0)
        b64raw = match.group(2)  # во второй группе у нас будет подстрока в base64
        raw = base64.b64decode(b64raw.encode())  # декодируем base64
        try:
            return raw.decode(encoding)  # пробуем перевести из байт в строку в той кодировке, которую узнали
        except Exception as ex:
            print(ex)
            return match.group(0)  # если не получилось, и словили ошибку, то вернем название кодировки
    """'Это очень сильное колдунство!' в общем, здесь мы сразу ищем в нашей строке закодированные в
    base64 куски, и к этим кускам применяем функцию replace, которая и формирует нам дешифрованную подстроку,
    на которую мы заменим зашифрованный кусок. И сразу возвращае строку с дешифрованными кусками base64!"""
    return re.sub(r'=\?(.*?)\?b\?(.*?)\?=', replace, value, flags=re.IGNORECASE)


def send(channel, command):
    """Функция для отправки сообщений"""
    channel.write(command)  # отправляем это сообщение
    channel.write('\n')  # добавляем знак переноса строки (ЭТ ВАЖНА!)
    channel.flush()  # очищаем буфер файл-подобного объекта сокета


def recv_line(channel):
    """Функция для приема однострочного ответа сервера"""
    response = channel.readline()[:-2]  # читаем ответ
    if response.startswith('-ERR'):
        raise POP3Exception(response)  # если сервер вернул нам сообщение об ошибке, то мы ее выводим


def recv_multiline(channel):
    """Функция, позволяющая считать многострочный ответ сервера"""
    recv_line(channel)  # сначала получаем кодовый ответ сервера
    lines = []  # заводим массивчик под линии сообщения
    while True:
        line = channel.readline()[:-2]  # читаем линию без знака переноса строки
        if line == '.':  # если достигли конца письма, то выходим
            break
        lines.append(line)  # заносим полученную линию в массив
    return lines  # возвращаем массивчик


def authentication(channel, username, password):
    """Функция для аутентификации пользователя"""
    send(channel, 'USER {}'.format(username))  # отправляем имя пользователя
    recv_line(channel)  # получаем ответ сервера
    channel.write('PASS {}\n'.format(password))  # отправляем пароль
    channel.flush()  # очищаем буфер файл-подобного объекта сокета
    recv_line(channel)  # получаем ответ сервера


def main(host='pop3.yandex.ru'):
    """
        Функция, возвращающая лист с информацией о письмах:
        От кого, кому, тема, дата, размер, инфа о вложениях
        """
    addr = socket.gethostbyname(host)  # получаем ip pop3-сервера
    sock = socket.socket()  # создаем udp-сокет
    sock.settimeout(10)  # устанавливаем таймаут
    sock = ssl.wrap_socket(sock)  # оборачиваем в ssl
    """Настраиваем порт. Для ssl-соединения используется 995 порт (для незащищенного - 110)"""
    port = 995
    sock.connect((addr, port))  # коннектимся по нужному адресу к нужному порту
    """Превращаем наш сокет в файл-подобный объект. Это довольно удобно для работы с сокетами"""
    channel = sock.makefile('rw', newline='\r\n', encoding='utf-8')
    recv_line(channel)  # получаем от сервера ответ после подключения
    username = input('Enter username: ')  # просим пользователя ввести адрес электронной почты
    password = input('Enter password: ')  # просим пользоватея ввести пароль
    authentication(channel, username, password)  # пробуем зайти по полученным данным
    letters_info = get_info(channel)
    print_list(letters_info)
    print_help()
    top_req_exp = re.compile('TOP (\d+) (\d+)')
    top_req_exp_cut = re.compile('TOP (\d+)')
    recv_req_exp = re.compile('RECV (\d+)')
    while True:
        command = input().upper()
        if command == 'EXIT':
            """заканчиваем работу программы"""
            print('Досвидос!')  # Прощаемся, выходим
            exit()
        elif command == 'HELP':
            print_help()
        elif command == 'LIST':
            print_list(letters_info)
        elif top_req_exp.match(command):
            info = top_req_exp.match(command)
            print_top(channel, info.group(1), info.group(2))
        elif top_req_exp_cut.match(command):
            info = top_req_exp_cut.match(command)
            print_top(channel, info.group(1), None)
        elif recv_req_exp.match(command):
            info = recv_req_exp.match(command)
            recv_letter(channel, int(info.group(1)), letters_info[int(info.group(1)) - 1])
        elif command == 'RECV ALL':
            recv_all(channel, letters_info)


if __name__ == '__main__':
    argparser = argparse.ArgumentParser()  # создаем парсер аргументов командной строки и добавляем аргументы
    argparser.add_argument('host',
                           help='Pop3 server address')  # адрес или имя pop3-сервера
    args = argparser.parse_args()  # парсим аргументы
    try:
        main(args.host)  # пробуем соединиться согласно ключам
        print('Successful')  # в случае успеха сообщаем об этом
    except KeyboardInterrupt:
        print('Досвидос!')  # прощаемся, если пользователь в процессе сам прекратил работу до завершения
    except Exception as error:
        print('Error:', error)  # принтуем, если вдруг какая ошибка
