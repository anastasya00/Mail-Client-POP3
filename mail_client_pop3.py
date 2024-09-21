import poplib
import smtplib
import getpass

class MailClientPOP3:

    providers = ['outlook']
    selected_provider = str()

    __login = str()
    __password = str()

    host = str()
    port = str()
    pop3 = None

    def __init__(self) -> None:
        self.greeting()

        if self.connect_to_pop3():
            print("TESTESTEST")
            self.pop3.quit()

    def greeting(self) -> None:
        print("Добро пожаловать в почтовый клиент POP3!\n")
        print("Каким почтовым провайдером вы пользуетесь?")
        for i, provider in enumerate(self.providers):   
            print(f'{i+1}. {provider}')

        choice = input("\nВведите номер провайдера: ")
        try:
            choice = int(choice)
            if 1 <= choice <= len(self.providers):
                self.selected_provider = str(self.providers[choice-1])
                print(f"Вы выбрали {self.selected_provider}.\n")

                self.__login = input("Логин: ")
                self.__password = getpass.getpass("Пароль: ")
                
        except ValueError:
            print("Введите число.")

    def connect_to_pop3(self):
        match self.selected_provider:
            case "outlook":
                self.host = "outlook.office365.com"
                self.port = "995"
            case _:
                print("Неизвестный провайдер.")
                return False
            
        try:
            self.pop3 = poplib.POP3_SSL(self.host, self.port)
            self.pop3.user(self.__login)
            self.pop3.pass_(self.__password)
            print("Подключение установлено с:\n")
            return True
        except Exception as e:
            print(f"Ошибка подключения: {e}.")
            return False


def main():
    MailClientPOP3()

if __name__ == "__main__":
    main()