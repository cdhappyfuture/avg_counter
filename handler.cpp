#include <ExcelFormat.h>
#include "handler.h"
#include "calc.h"
#include <cstdio>

string DaysToDate(int days)
{
    char cs[20];
    sprintf(cs, "%d", days);
    string s = cs;
    return s;
}

bool FileExist(string fileName)
{
    std::ifstream file(fileName.c_str());
    return file.good();
}

bool Handler(char* jour_file)
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
           day_file;

    while (row != j_sheet->GetTotalRows()-1)
    {//Обрабатываем каждый диапазон
        if (beg_cell->GetDouble() == 0)
        {
            date_cell = j_sheet->Cell(row, 0);
            if (date_cell->GetInteger())
                cur_date = DaysToDate(date_cell->GetInteger());
            else
                break;
            day_file = cur_date + ".xls";
        }
        if (FileExist(day_file))
        {
            Calc::Result* result = Calc::calc(beg_cell->GetDouble(),end_cell->GetDouble(),day_file); // получаем результаты вычислений
            if (result)
            {
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
        row++;
        end_cell = j_sheet->Cell(row, END_CI);
        beg_cell = j_sheet->Cell(row, BEG_CI);
    }
    journal.SaveAs("newjour.xls");
    return 0;
}
