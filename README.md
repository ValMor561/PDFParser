Необходимо: установленный python
Открываем консоль и запускаем виртуальное окружение командой
```shell
python -m venv venv
```
Активируем его
Windows
```shell
.\venv\Scripts\Activate.ps1
```
Linux
```bash
source venv/bin/activate
```
Устанавливаем все необходимые библиотеки
```shell
pip install -r requirements.txt
```
Запускаем программу
```shell
python main.py <pdf_filename> <xlxs_filename>
```
Пример
```shell
python main.py test.pdf test.xlxs
```
Для удобства пользования можно вызвать справку
```shell
python main.py --help
```

