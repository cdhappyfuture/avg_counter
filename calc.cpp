#include <ExcelFormat.h>
#include "calc.h"
#include <iostream>

namespace Calc
{
    class CellError
    {};

    double TimeConvert(string s)
        // Принимает время вида "13:04" или " 9:23" и возвращает эквивалент
        // в double для дальнейшего сравнения с временем из журнала
    {
        int i =0,hour =0,min =0;
        hour = s[1]-'0';
        min = (s[3]-'0')*10 + s[4]-'0';
        if (s[0] != ' ')
            hour += (s[0]-'0')*10;
        return (hour*60+min)/(24*60.0);
    }

    double calc_UNI(int beg_RI, int end_RI, int beg_CI, int end_CI, ExcelFormat::BasicExcelWorksheet* sheet)
    {
        ExcelFormat::BasicExcelCell *cell;
        double sum = 0;
        int count = 0;
        for (int r = beg_RI; r <= end_RI; r++)
            for (int c = beg_CI; c <= end_CI; c++)
            {
                cell = sheet->Cell(r,c);
                if (cell->GetDouble())
                    sum += cell->GetDouble();
                else
                    sum += cell->GetInteger();
                //    throw CellError(;
                count++;
            }
        return sum/count;
    }
}
using namespace Calc;

Result* Calc::calc(double beg_time, double end_time, string day_file)
{
    const int TIME_CI = 3; // Индекс колонки с временем
    //Получаем лист из файла
    ExcelFormat::BasicExcel day_book;
    day_book.Load(day_file.c_str());
    ExcelFormat::BasicExcelWorksheet* sheet = day_book.GetWorksheet(0);

    //Ищем диапазон строк, входящих в период
    int beg, end, row = 1;
    ExcelFormat::BasicExcelCell* cell_time = sheet->Cell(row, TIME_CI); // первая запись в таблице
    if (TimeConvert(cell_time->GetString()) > end_time)
        return NULL;
    while (TimeConvert(cell_time->GetString()) < beg_time)
    {
        cell_time = sheet->Cell(++row, TIME_CI);
        if (row == sheet->GetTotalRows()-1)
            return NULL;
    }
    beg = row;
    while (TimeConvert(cell_time->GetString()) < end_time)
    {
        cell_time = sheet->Cell(++row, TIME_CI);
        if (row == sheet->GetTotalRows()-1)
            return NULL;
    }
    end = ++row;

    // Вычисляем средние значения
    Result *res = new Result;
    try
    {
        res->E = calc_UNI(beg, end, E, E, sheet);
        res->F = calc_UNI(beg, end, F, F, sheet);
        res->GH =  calc_UNI(beg, end, G, H, sheet);
        res->IL = calc_UNI(beg, end, I, L, sheet);
        res->MP = calc_UNI(beg, end, M, P, sheet);
        res->QT = calc_UNI(beg, end, Q, T, sheet);
    }
    catch (CellError err)
    {
        return NULL;
    }
    return res;
}
