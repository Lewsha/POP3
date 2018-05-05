import socket
import ssl
import argparse
import base64
import re


class Pop3Exception(Exception):
    """Наша собственная ошибка, ееее!"""
    def __init__(self, msg):
        super().__init__(msg)  # Тупо super'имся от родителя с нужным сообщением


class MailStruct:
    """Класса объета письма (данные о письме, не само письмо)"""
    def __init__(self, mail_from, mail_to, mail_subject,
                 mail_date, mail_size, mail_attachments):
        self.mail_from = mail_from  # от кого письмо
        self.mail_to = mail_to  # кому письмо
        self.mail_subject = mail_subject  # тема письма
        self.mail_date = mail_date  # дата отправки
        self.mail_size = mail_size  # размер письма
        self.mail_attachments = mail_attachments  # данные о вложениях письма

    def __repr__(self):
        """Переопределяем представление объекта"""
        headers = ('From: {}\n'
                   'To: {}\n'
                   'Subject: {}\n'
                   'Date: {}\n'
                   'Mail Size: {} bytes\n'
                   'Attachments list:\n')\
            .format(self.mail_from, self.mail_to,
                    self.mail_subject, self.mail_date,
                    self.mail_size)  # вставляем нужные данные в нужные места

        attachments = []  # создаем лист под вложения
        for idx, (file_name, file_size) in enumerate(self.mail_attachments):
            attachments.append(
                '{:3d}) FileName: {:s}\n     FileSize: {} bytes'.format(
                    idx + 1, file_name or 'none', file_size or 'none'))  # в цикле добавляем инфу о вложениях
        return headers + ('\n'.join(attachments) if attachments else 'None')  # возвращаем грамотно
        # скомпанованное представление


def recv_line(handle):
    """Функция для приема однострочного ответа сервера"""
    response = handle.readline()[:-2]  # читаем ответ
    if response.startswith('-ERR'):
        raise Pop3Exception(response)  # если сервер вернул нам сообщение об ошибке, то мы ее выводим


def recv_multiline(handle):
    """Функция, позволяющая считать многострочный ответ сервера"""
    recv_line(handle)  # сначала получаем ответ сервера
    lines = []  # заводим массивчик под линии сообщения
    count = 0  # заводим счетчик
    while True:
        line = handle.readline()[:-2]  # читаем линию
        if count < 5:  # если лимит не превышен, заносим в логи
            count += 1  # увеличиваем счетчик
        if count == 5:  # если дошли до предела, логируем, что еще не все
            count += 1  # увеличиваем счетчиа
        if line == '.':  # если достигли конца письма, то выходим
            break
        lines.append(line)  # заносим полученную линию в массив
    return lines  # возвращаем массивчик


def send(handle, command):
    """Функция для отправки сообщений"""
    handle.write(command)  # отправляем это сообщение
    handle.write('\n')  # добавляем знак переноса строки (ЭТ ВАЖНА!)
    handle.flush()  # очищаем буфер файл-подобного объекта сокета


def decode_inline_base64(s):
    """Функция декодирования base64"""
    def replace(match):
        encoding = match.group(1).lower()
        b64raw = match.group(2)
        raw = base64.b64decode(b64raw.encode())
        try:
            return raw.decode(encoding)
        except Exception as ex:
            print(ex)
            return match.group(0)
    return re.sub(r'=\?(.*?)\?b\?(.*?)\?=', replace, s, flags=re.IGNORECASE)


def find_attachments(message):
    flags = re.IGNORECASE | re.MULTILINE | re.DOTALL
    attach_regexp = r'^Content-Disposition:\sattachment(.*?)(?:^\w|\Z|\n\n)'
    result = []
    for match in re.finditer(attach_regexp, message, flags):
        file_name = None
        file_size = None
        attach_header = regexp_post_processing(match.group(1))
        name_match = re.search(
            r'filename="(.*?)"', attach_header, re.IGNORECASE)
        if name_match:
            file_name = decode_inline_base64(name_match.group(1))
        size_match = re.search(
            r'size=(\d+)', attach_header, re.IGNORECASE)
        if size_match:
            file_size = int(size_match.group(1))
        else:
            try:
                content_start_pos = message.find('\n\n', match.span()[0]) + 2
                content_end_pos = message.find('\n\n', content_start_pos)
                content = message[content_start_pos:content_end_pos]
                raw = re.sub(r'[^A-Za-z0-9+/=]', '', content).encode()
                file_size = len(base64.b64decode(raw))
            except Exception as ex:
                print(ex)
        result.append((file_name, file_size))
    return result


def regexp_post_processing(s):
    """Функция для, скажем так, улучшения читаемости получаемых заголовков письма"""
    flags = re.IGNORECASE | re.MULTILINE | re.DOTALL  # выставляем флаги для регулярного выражения
    return re.sub(r'^\s*', '', s.strip(), flags=flags).replace('\n', '')  # возвращаем, по сути, то, что получили


def find_header(header, headers):
    """Функция для вытяжки из письма заголовков"""
    flags = re.IGNORECASE | re.MULTILINE | re.DOTALL  # устанавливаем флаги для регулярного выражения
    match = re.search(r'^{}:(.*?)(?:^\w|\Z|\n\n)'
                      .format(header), headers, flags)  # ищем заголовки в письме
    if match:
        return regexp_post_processing(match.group(1))  # если нашли, возвращаем грамотно скомпанованные заголовки


def parse_message(message):
    """
    Extract important headers and information from mail
    """
    important_headers = ['From', 'To', 'Subject', 'Date']
    important_headers_values = []
    headers, *body = message.split('\n\n')
    for header in important_headers:
        important_headers_values.append(
            find_header(header, headers) or 'unknown')
    try:
        attachments = find_attachments(message)
    except Exception as ex:
        attachments = []
    important_headers_values = list(
        map(decode_inline_base64, important_headers_values))
    important_headers_values.append(attachments)
    return important_headers_values


def pop3_auth(handle, username, password):
    """Функция для аутентификации пользователя"""
    send(handle, 'USER {}'.format(username))  # отправляем имя пользователя
    recv_line(handle)  # получаем ответ сервера
    handle.write('PASS {}\n'.format(password))  # отправляем пароль
    handle.flush()  # очищаем буфер файл-подобного объекта сокета
    recv_line(handle)  # получаем ответ сервера


def pop3_retr_mail(handle, msg_id):
    """Функция, скачивающая письмо с сервера"""
    send(handle, 'RETR {}'.format(msg_id))  # посылаем команду RETR с номером нужного письма
    return '\n'.join(recv_multiline(handle))  # возвращаем полученное сообщение


def pop3_get_summary_info(host):
    """
    Функция, возвращающая лист с информацией о письмах:
    От кого, кому, тема, дата, размер, инфа о вложениях
    """
    addr = socket.gethostbyname(host)  # получаем ip pop3-сервера
    sock = socket.socket()  # создаем udp-сокет
    sock.settimeout(10)  # устанавливаем наш таймаут
    sock = ssl.wrap_socket(sock)  # если требуется защищенное соединение, оборачиваем в ssl
    port = 995  # выбираем порт в зависимости от типа соединения
    sock.connect((addr, port))  # коннектимся по нужному адресу к нужному порту
    """Превращаем наш сокет в файл-подобный объект. Это довольно удобно для работы с сокетами"""
    handle = sock.makefile('rw', newline='\r\n', encoding='utf-8')
    recv_line(handle)  # получаем от сервера ответ после подключения
    username = input('Enter username: ')  # просим пользователя ввести адрес электронной почты
    password = input('Enter password: ')  # просим пользоватея ввести пароль
    pop3_auth(handle, username, password)  # пробуем авторизоваться по полученным данным
    message_structs = []  # создаем лист под инфу о письмах
    send(handle, 'LIST')  # посылаем команду LIST для получения списка писем
    """Из полученного ответа вынимаем интересующую информацию"""
    for line in recv_multiline(handle):
        msg_id, msg_size = map(int, line.split())  # из строки забираем номер письма и его размер
        message = pop3_retr_mail(handle, msg_id)  # скачиваем само письмо
        msg_from, msg_to, msg_subj, msg_date, msg_attachments \
            = parse_message(message)  # парсим сообщение и забираем нужные данные о письме
        message_structs.append(MailStruct(
            msg_from, msg_to, msg_subj, msg_date, msg_size, msg_attachments))  # создаем структуру из полученных данных
    send(handle, 'QUIT')  # закрываем соединение с сервером
    recv_line(handle)  # получаем ответ
    return message_structs  # возвращаем полученные данные


def main():
    """Главная функция программы, запускающая весь процесс"""
    argparser = argparse.ArgumentParser()  # создаем парсер аргументов командной строки и добавляем аргументы
    argparser.add_argument('host',
                           help='Pop3 server address')  # адрес или имя pop3-сервера
    args = argparser.parse_args()  # парсим аргументы
    try:
        messages = pop3_get_summary_info(args.host)  # пробуем получить общую инфу о письмах
    except Exception as ex:
        print(ex)  # если ловим ошибку, то принтуем сообщение и выходим
        return
    except KeyboardInterrupt:
        print('Досвидос!')  # если пользователь прервал работу программы до окончания ее работы, то прощаемся и выходим
        return
    for idx, message in enumerate(messages):
        print('##### Message {:4d} #####'.format(idx + 1))  # принтуем полученную инфу на экран
        try:
            print(message)
        except UnicodeEncodeError:
            print('Can\'t print this')  # если что-то не вышло с декодировкой, то принтуем инфу об этом
        print()  # сообщения разделяем пустыми строками


if __name__ == '__main__':
    main()  # запускаем прогу
