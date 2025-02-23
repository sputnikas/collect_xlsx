﻿#pragma once
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Задача:
// Из множества файлов xlsx с одинаковой структурой получить один в котором будут 
// 1. Один общий столбец заголовков
// 2. Остальные столбцы берутся из файлов, вверху название файла или какая-то определнная ячейка
// 3. Должен быть интерфейс
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include <iostream>
#include <filesystem>
#include <string>
#include <vector>
#include <regex>
#include <stdexcept>
#include <exception>
#include <cstdarg>

#include <xlnt/xlnt.hpp>

using namespace std;
namespace fs = filesystem;

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Обработка ошибок
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

struct Error : exception {
    char text[1000];
    ////////////////////////////////////////////////////////////////////////////////////
    Error(char const *fmt, ...);
    ////////////////////////////////////////////////////////////////////////////////////
    char const *what() const throw() { return text; }
};

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Функции для работы с прямоугольными областями в таблице
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

struct MRange {
    int begin_row;
    int begin_col;
    int end_row;
    int end_col;
    ////////////////////////////////////////////////////////////////////////////////////
    MRange(const string &range);
};

bool is_range(const string &range);
int calc_column(const string &col);
pair<int, int> calc_cell(const string &addr);
void copy_value(xlnt::worksheet &outws, int ocol, int orow, xlnt::worksheet &ws, int icol, int irow);

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Сборка одной таблицы из нескольких
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

struct CollectXLSX {
    // Входные параметры
    string input_dir;
    string output_name;
    int sheet_index;
    string range_ref;
    string range_1ref;
    // Промежуточные параметры
    vector<string> input_xlsx; // хранит имена файлов в директории
    vector<bool> collected;
    int last_col;
    xlnt::workbook outwb;
    xlnt::worksheet outws;
    bool created_input;
    ////////////////////////////////////////////////////////////////////////////////////
    CollectXLSX();
    ////////////////////////////////////////////////////////////////////////////////////
    void add_xlsx_column(const string &xlsx_name);
    int get_1index_collected();
    void add_1xlsx_column();
    void create_input_xlsx();
    void collect();
    void clear();
};

