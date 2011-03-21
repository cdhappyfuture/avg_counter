#include <iostream>
#include "handler.h"

int main(int argc, char* argv[])
{
    if (argc != 3)
    {
        std::cout << "Первым аргументом должно быть имя файла - журнала, а вторым шаблон файлов с замерами содержащий DD, MM и YYYY вместо настоящих чисел, например UF_804001_YYYY_MM_DD.xls\n";
        return 1;
    }
    if (FileExist(argv[1]))
    {
        if (Handler(argv[1],argv[2]))
        {
            std::cout << "Ошибка в программе. Завершение\n";
            return 2;
        }
    }
    else
    {
        std::cout << "Отсутствует файл журнала\n";
        return 3;
    }
    std::cout << "Программа завершила работу без ошибок\n";
    return 0;
}
