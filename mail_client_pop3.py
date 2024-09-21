import poplib
import smtplib
import getpass
import email
from email.header import decode_header
from email.mime.text import MIMEText

class MailClientPOP3:

    __login = str()
    __password = str()

    providers = ['outlook']
    selected_provider = str()

    actions = ['Просмотреть список писем', 'Отправить письмо', 'Удалить письмо', 'Выйти']

    pop_host = str()
    pop_port = str()
    smtp_host = str()
    smtp_port = str()

    mailbox = None

    def __init__(self) -> None:
        self.greeting()

        if self.connect_to_pop3():
            self.menu()

    def greeting(self) -> None:
        print("Добро пожаловать в почтовый клиент POP3!\n")
        print("Каким почтовым провайдером вы пользуетесь?")
        for i, provider in enumerate(self.providers):   
            print(f'{i+1}. {provider}')

        while True:
            choice = input("\nВведите номер провайдера: ")
            try:
                choice = int(choice)
                if 1 <= choice <= len(self.providers):
                    self.selected_provider = str(self.providers[choice-1])
                    print(f"Вы выбрали {self.selected_provider}.\n")

                    self.__login = input("Логин: ")
                    self.__password = getpass.getpass("Пароль: ")
                    break
                else:
                    print(f"Неверный выбор. Попробуйте снова.")                    
            except ValueError:
                print("Введите число.")

    def connect_to_pop3(self) -> bool:
        match self.selected_provider:
            case "outlook":
                self.pop_host = "outlook.office365.com"
                self.pop_port = "995"
                self.smtp_host = "smtp-mail.outlook.com"
                self.smtp_port = "587"
            case _:
                print("Неизвестный провайдер.")
                return False
            
        try:
            self.mailbox = poplib.POP3_SSL(self.pop_host, self.pop_port)
            self.mailbox.user(self.__login)
            self.mailbox.pass_(self.__password)
            print("Подключение установлено с:\n")
            return True
        except Exception as e:
            print(f"Ошибка подключения: {e}.")
            return False

    def menu(self) -> None:
        while True:
            print("Меню:")
            for i, action in enumerate(self.actions):
                print(f"{i+1}. {action}")

            choice = input("\nВыберите номер действия: ")
            try:
                choice = int(choice)
                if 1 <= choice <= len(self.actions):
                    
                    match choice:
                        case 1:
                            self.view_list_emails()
                        case 2:
                            if self.send_email():
                                print("Письмо успешно отправлено.\n")
                        case 3:
                            if self.delete_email():
                                print("Письмо успешно удалено.\n")
                        case 4:
                            self.__delattr__()
                            print("Вы вышли из почтового клиента.")
                            break
                else:
                    print(f"Неверный выбор. Попробуйте снова.")
            except ValueError:
                print("Введите число.")
    
    def view_list_emails(self) -> None:
        try:
            emails_count, _ = self.mailbox.stat()
            print(f"Количество писем: {emails_count}")
            
            messagers = self.mailbox.list()[1]
            for i in range(len(messagers)):
                raw_email = b"\n".join(self.mailbox.retr(i+1)[1])
                parsed_email = email.message_from_bytes(raw_email)

                print(f'\n--------------- Письмо №{i+1} ---------------')
                print('От кого:', self.decoding(parsed_email['From']))
                print('Кому:', parsed_email['to'])
                print('Дата:', parsed_email['Date'])
                print('Тема: ', self.decoding(parsed_email['Subject']))
                print('--------------------------------------------\n')

        except Exception as e:
            print("Ошибка при получении списка писем.")

    def decoding(self, text):
        decoded_text = decode_header(text)[0]
        if decoded_text[1] is not None:
            return decoded_text[0].decode(decoded_text[1])
        else:
            return decoded_text[0].decode('utf-8', 'ignore')

    def delete_email(self) -> bool:
        email_number = int(input("Введите номер письма для удаления: "))

        try:
            if 1 <= email_number <= self.mailbox.stat()[0]:
                self.mailbox.dele(email_number)
                return True
            else:
                print("Неверный номер письма.")
                return False
            
        except Exception as e:
            print(f"Ошибка при удалении письма: {e}")

    def send_email(self) -> bool:
        try:
            from_addr = self.__login
            to_addr = input("Кому: ")
            subject = input("Тема письма: ")
            text = input("Содержимое письма: ")

            msg = MIMEText(text)
            msg['Subject'] = subject
            msg['From'] = from_addr
            msg['To'] = to_addr

            send = smtplib.SMTP(self.smtp_host, self.smtp_port)
            send.ehlo()
            send.starttls()
            send.login(self.__login, self.__password)
            send.sendmail(from_addr, [to_addr], msg.as_string())
            send.quit()

            return True
        
        except Exception as e:
            print(f"Ошибка при отправке письма: {e}")
            return False

    def __delattr__(self) -> None:
        self.mailbox.quit()


def main():
    MailClientPOP3()

if __name__ == "__main__":
    main()