#include <iostream>
#include "handler.h"

int main(int argc, char* argv[])
{
    if (argc != 2)
    {
        std::cout << "Первым аргументом должно быть имя файла - журнала\n";
        return 1;
    }
    if (Handler(argv[1]))
    {
        std::cout << "Ошибка в программе. Завершение\n";
        return 1;
    }
    std::cout << "Программа завершила работу без ошибок\n";
    return 0;
}
