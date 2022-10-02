# Преобразование Exel-файлов из формата xls в xlsx средствами Windows

from win32com import client as wc
import os


def xls_to_xlsx(files_path, filename_template='*', isrm_file=True):
    # функция преобразования файлов xls в файлы xlsx
    # использует запуск как Windows Service
    # files_path - путь к целевой папке, в которой будут преобразованы файлы, включая вложенные папки
    # filename_template - шаблон для имени файла, указывает на то, какие файлы нужно преобразовывать
    # необходимо указать имя файла и тогда преобразование будет только для файлов с этим именем
    # если указать '*' - то преобразованы будут все файлы формата xls в целевой папке
    # по умолчанию имеет зничение '*'
    # isrm_file - индикатор, указывает удалять ли исходный файл после преобразования
    # True - исходный файл будет удален после преобразования
    # False - исходный файл не буде удаляться после преобразования
    # по умолчанию имеет значение True

    # взводим индикатор отсутствия ошибки
    no_error_flag = True

    # устанавливаем связь с приложением Microsoft Excel
    try:
        e = wc.Dispatch("Excel.Application")
        e.Visible = 0
        e. DisplayAlerts = 0
    except:
        no_error_flag = False
        print('Ошибка связи с приложением Microsoft Excel')
        e.Quit()  # разрываем связь с приложением Microsoft Excel
        return no_error_flag

    # формируем перечень файлов для извлечения данных
    try:
        paths = []

        for root, dirs, files in os.walk(files_path):
            for file in files:
                if filename_template == '*':
                    if file.endswith('xls') and not file.startswith('~'):
                        paths.append(os.path.join(root, file))
                else:
                    if file.endswith('xls') and not file.startswith('~') and (os.path.splitext(file)[0]==filename_template):
                        paths.append(os.path.join(root, file))
    except:
        no_error_flag = False
        print('Ошибка при поиске файлов .xls')
        e.Quit()                                                  # разрываем связь с приложением Microsoft Excel
        return no_error_flag

    # преобразуем формат xls в формат xlsx
    try:
        if len(paths)!=0:
            for path in paths:
                doc = e.Workbooks.Open(os.path.abspath(path))
                doc.SaveAs(os.path.abspath(path)+'x', FileFormat = 51)  # 51 - формат файла Open XML Workbook (xlsx)
                                                                        # 56 - Excel 97-2003 Workbook (xls)
                doc.Close(SaveChanges=0)

                if isrm_file:
                    os.remove(path)                               # удаляем исходный файл после преобразования
        else:
            no_error_flag = False
            print('Файлы для преобразования не найдены')
            e.Quit()                                              # разрываем связь с приложением Microsoft Excel
            return no_error_flag

    except:
        no_error_flag = False
        print('Ошибка при преобразовании .xls в .xlsx')
        e.Quit()                                      # разрываем связь с приложением Microsoft Excel
        return no_error_flag

    return no_error_flag