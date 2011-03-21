#include <ExcelFormat.h>
#include "handler.h"
#include "calc.h"
#include <cstdio>
#include <sstream>

struct Date 
{
    int y;
    int m;
    int d;
};

Date* DaysToDate(int days)
{//принимает количество дней и возвращает дату в виде строки
    Date* date = new Date;
    date->m = 0;
    const int BEGIN = 1900;
    int dm[][12] = {{31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31},
        {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}};
    date->y = days*4 / (static_cast<int>(365.25*4));                                         //количество лет в полученом периоде
    int usegod = (date->y % 4 ? 0 : 1);                               //если год вясокосный то 0 иначе 1
    date->d = days - date->y*365 - date->y/4 - usegod;                        //количество дней в неполном году
    while (date->d > dm[usegod][date->m])                               //пока полный месяц
    {
        date->d -= dm[usegod][date->m++];                                 //вычитаем его длину из неполного года                                              //переходим к следующему месяцу
    }
    date->m++;
    date->y += BEGIN;
    return date;
}

string NameFromTemplate(string templ, Date* date)
{

    ostringstream s;
    int i = templ.find("DD");
    if (date->d <= 9)
        s << '0' << date->d;
    else
        s << date->d;
    templ.replace(i, 2, s.str());
    i = 0;
    i = templ.find("MM");
    s.str("");
    if (date->m <= 9)
        s << '0' << date->m;
    else
        s << date->m;
    templ.replace(i, 2, s.str());
    i = templ.find("YYYY");
    s.str("");
    s << date->y;
    templ.replace(i, 4, s.str());
    return templ;
}

bool FileExist(char* fileName)
{
    std::ifstream file(fileName);
    return file.good();
}

bool Handler(char* jour_file, char* tem)
{
    const int BEG_CI = 5; // Индекс столбца с началом периода
    const int END_CI = 6; // Индекс столбца с конецом периода
    string journal_file = jour_file;

    ExcelFormat::BasicExcel journal;
    journal.Load(journal_file.c_str());
    ExcelFormat::BasicExcelWorksheet* j_sheet = journal.GetWorksheet(0);
    
    //Рассчитываем индексы столбцов добавляемых значений
    int E_CI = j_sheet->GetTotalCols(); // Результат по столбцу E будет помещен в первый пустой столбец журнала
    int F_CI = E_CI + 1; // Индекс столбца со средним по F
    int GH_CI = E_CI + 2; // Индекс столбца со средним по G
    int IL_CI = E_CI + 3;//Индекс столбца со средним из столбцов C-I
    int MP_CI = E_CI + 4;
    int QT_CI = E_CI + 5;

    ExcelFormat::BasicExcelCell *beg_cell, *end_cell, *date_cell, *cell;
    int row = 2; // Индекс строки журнала, где лежит первая дата
    
    end_cell = j_sheet->Cell(row, END_CI);
    beg_cell = j_sheet->Cell(row, BEG_CI);
    string cur_date,
           day_file,
           templ = tem;
    while (row != j_sheet->GetTotalRows())
    {//Обрабатываем каждый диапазон
        if (beg_cell->GetDouble() == 0)
        {//Если блок диапазонов кончился, смотрим новую дату
            date_cell = j_sheet->Cell(row, 0);
            if (date_cell->GetInteger())
            {
                day_file = NameFromTemplate(templ, DaysToDate(date_cell->GetInteger())) ;
            }
            else
                break;
        }
        if (FileExist((char*)day_file.c_str()))
        {//Если есть файл на текущую дату, обрабатываем его
            Calc::Result* result = Calc::calc(beg_cell->GetDouble(),end_cell->GetDouble(),day_file); // получаем результаты вычислений
            if (result)
            { //Если не произошло ошибок во время вычислений, заполняем журнал подсчитанными значениями
                cell = j_sheet->Cell(row, E_CI);
                cell->Set(result->E);
                cell = j_sheet->Cell(row, F_CI);
                cell->Set(result->F);
                cell = j_sheet->Cell(row, GH_CI);
                cell->Set(result->GH);
                cell = j_sheet->Cell(row, IL_CI);
                cell->Set(result->IL);
                cell = j_sheet->Cell(row, MP_CI);
                cell->Set(result->MP);
                cell = j_sheet->Cell(row, QT_CI);
                cell->Set(result->QT);
            }
            delete result;
        }
        else
            std::cout << "Файл " << day_file << " отсутствует в каталоге\n";
        row++;
        end_cell = j_sheet->Cell(row, END_CI);
        beg_cell = j_sheet->Cell(row, BEG_CI);
    }
    string res_file = "new";
    res_file += jour_file;
    journal.SaveAs(res_file.c_str());
    return 0;
}
