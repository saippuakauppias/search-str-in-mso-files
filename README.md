search-str-in-mso-files
======================

Небольшой скрипт для поиска заданной строки в *.docx и *.xlsx файлах.


Установка
---------

По-хорошему нужно поставить виртуальное окружение:

    Установить pip: http://pypi.python.org/pypi/pip (http://www.pip-installer.org/en/latest/installing.html#using-the-installer)
    Установить virtualenv: http://pypi.python.org/pypi/virtualenv
    Установить virtualenvwrapper: http://pypi.python.org/pypi/virtualenvwrapper
    Настроить виртуальное окружение (в ~/.bashrc):
        export WORKON_HOME=$HOME/.virtualenvs
        export PROJECT_HOME=$HOME/projects
        export VIRTUALENVWRAPPER_VIRTUALENV_ARGS='--no-site-packages'
        export VIRTUALENVWRAPPER_LOG_DIR=$WORKON_HOME
        source "/usr/local/bin/virtualenvwrapper.sh"
        export PIP_REQUIRE_VIRTUALENV=true
    Создать виртуальное окружение: $ mkproject testproject
    Скопировать в ~/projects/testproject содержимое архива
    Активировать виртуальное окружение: $ workon testproject
    Установить зависимости: $ pip install -r requirements.txt

По-простому (но я бы не советовал засорять глобальный site-packages питона):

    Установить pip: http://pypi.python.org/pypi/pip (http://www.pip-installer.org/en/latest/installing.html#using-the-installer)
    Скопировать куда-нибудь содержимое архива
    Установить зависимости: $ pip install -r requirements.txt

В зависимостях тянется lxml - пакет для работы с xml.
Для его установки нужен gcc (а тот в свою очередь, по-моему, требует python-dev), libxml2-dev и libxslt-dev.


Тестирование
------------

Протестировано на Python 2.7.2 (Mac OS X 10.8.2) и Python 2.7.3 (Ubuntu 12.04).

Для запуска тестов нужно выполнить из командной строки:

    $ python -m unittest tests


Поиск файлов
------------

Для запуска поиска файлов нужно выполнить из командной строки:

    $ python search_str/find_in_dir.py -d 'директория' -s 'строка'

Например, можно поискать в фикстурах, которые прилагаются к тестам:

    $ python search_str/find_in_dir.py -d 'search_str/fixtures/' -s 'Русский'


[![Bitdeli Badge](https://d2weczhvl823v0.cloudfront.net/saippuakauppias/search-str-in-mso-files/trend.png)](https://bitdeli.com/free "Bitdeli Badge")

