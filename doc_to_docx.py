from win32com import client as wc
import os


def doc_to_docx(files_path, filename_template='*', isrm_file=True):
    # функция преобразования файлов doc в файлы docx
    # использует запуск как Windows Service
    # files_path - путь к целевой папке, в которой будут преобразованы файлы, включая вложенные папки
    # filename_template - шаблон для имени файла, указывает на то, какие файлы нужно преобразовывать
    # необходимо указать имя файла и тогда преобразование будет только для файлов с этим именем
    # если указать '*' - то преобразованы будут все файлы формата doc в целевой папке
    # по умолчанию имеет зничение '*'
    # isrm_file - индикатор, указывает удалять ли исходный файл после преобразования
    # True - исходный файл будет удален после преобразования
    # False - исходный файл не буде удаляться после преобразования
    # по умолчанию имеет значение True

    # взводим индикатор отсутствия ошибки
    no_error_flag = True

    # устанавливаем связь с приложением Microsoft Word
    try:
        w = wc.Dispatch('Word.Application')
    except:
        no_error_flag = False
        print('Ошибка связи с приложением Microsoft Word')
        w.Quit()  # разрываем связь с приложением Microsoft Word
        return no_error_flag

    # формируем перечень файлов для извлечения данных
    try:
        paths = []

        for root, dirs, files in os.walk(files_path):
            for file in files:
                if filename_template == '*':
                    if file.endswith('doc') and not file.startswith('~'):
                        paths.append(os.path.join(root, file))
                else:
                    if file.endswith('doc') and not file.startswith('~') and (os.path.splitext(file)[0]==filename_template):
                        paths.append(os.path.join(root, file))
    except:
        no_error_flag = False
        print('Ошибка при поиске файлов .doc')
        w.Quit()                                                  # разрываем связь с приложением Microsoft Word
        return no_error_flag

    # преобразуем формат doc в формат docx
    try:
        if len(paths)!=0:
            for path in paths:
                doc = w.Documents.Open(os.path.abspath(path))
                doc.SaveAs(os.path.abspath(path)+'x',16)
                doc.Close()

                if isrm_file:
                    os.remove(path)                               # удаляем исходный файл после преобразования
        else:
            no_error_flag = False
            print('Файлы для преобразования не найдены')
            w.Quit()                                              # разрываем связь с приложением Microsoft Word
            return no_error_flag

    except:
        no_error_flag = False
        print('Ошибка при преобразовании .doc в .docx')
        w.Quit()                                                  # разрываем связь с приложением Microsoft Word
        return no_error_flag

    return no_error_flag