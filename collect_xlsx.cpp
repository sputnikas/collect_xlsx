#include "collect_xlsx.h"

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ќбработка ошибок
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Error::Error(char const *fmt, ...) {
    va_list ap;
    va_start(ap, fmt);
    vsnprintf(text, sizeof(text), fmt, ap);
    va_end(ap);
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ‘ункции дл¤ работы с пр¤моугольными област¤ми в таблице
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


bool is_range(const string &range) {
    const regex txt_re("[A-Z]+\\d+(:[A-Z]+\\d+)?");
    smatch m;
    return regex_match(range, m, txt_re);
}

int calc_column(const string &col) {
    int icol = 0;
    for (char a : col) {
        icol = icol * 26 + (a - 'A' + 1);
    }
    return icol;
}

pair<int, int> calc_cell(const string &addr) {
    const regex txt_re("([A-Z]+)(\\d+)");
    smatch m;
    if (regex_match(addr, m, txt_re)) {
        return {calc_column(m[1]), atoi(m[2].str().c_str())};
    } else {
        throw Error("Couldn't calculate row and column for address = %s", addr.c_str());
    }
}

MRange::MRange(const string &range) {
    if (is_range(range)) {
        size_t ddoti = range.find(':');
        if (ddoti != string::npos) {
            string c1 = range.substr(0, ddoti);
            string c2 = range.substr(ddoti + 1);
            auto ic1 = calc_cell(c1);
            auto ic2 = calc_cell(c2);
            begin_row = min(ic1.second, ic2.second);
            begin_col = min(ic1.first, ic2.first);
            end_row = max(ic1.second, ic2.second);
            end_col = max(ic1.first, ic2.first);
        } else {
            auto ic = calc_cell(range);
            begin_row = ic.second;
            begin_col = ic.first;
            end_row = ic.second;
            end_col = ic.first;
        }
    } else {
        begin_row = 0;
        begin_col = 0;
        end_row = 0;
        end_col = 0;
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// —борка одной таблицы из нескольких
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void copy_value(xlnt::worksheet &outws, int ocol, int orow, xlnt::worksheet &ws, int icol, int irow) {
    switch (ws.cell(icol, irow).data_type()) {
        case xlnt::cell_type::boolean:
            // printf("boolean\n");
            outws.cell(ocol, orow).value(ws.cell(icol, irow));
            break;
        case xlnt::cell_type::date:
            // printf("date\n");
            outws.cell(ocol, orow).value(ws.cell(icol, irow));
            break;
        case xlnt::cell_type::empty:
            // printf("empty\n");
            break;
        case xlnt::cell_type::error:
            // printf("error\n");
            break;
        case xlnt::cell_type::formula_string:
            // printf("formula_string\n");
            outws.cell(ocol, orow).value(ws.cell(icol, irow));
            break;
        case xlnt::cell_type::inline_string:
            // printf("inline_string\n");
            outws.cell(ocol, orow).value(ws.cell(icol, irow).to_string());
            break;
        case xlnt::cell_type::number:
            // printf("number\n");
            outws.cell(ocol, orow).value(ws.cell(icol, irow));
            break;
        case xlnt::cell_type::shared_string:
            // printf("shared_string\n");
            outws.cell(ocol, orow).value(ws.cell(icol, irow).to_string());
            break;
    }
    outws.cell(ocol, orow).data_type(ws.cell(icol, irow).data_type());
}

CollectXLSX::CollectXLSX() {
    input_dir = "./";
    output_name = "1.xlsx";
    sheet_index = 0;
    range_ref = "B2:B20";
    range_1ref = "A2:A20";
    last_col = 1;
    outws = outwb.active_sheet();
    outws.title("Data");
    created_input = false;
}

void CollectXLSX::add_xlsx_column(const string &xlsx_name) {
    xlnt::workbook wb;
    wb.load(xlsx_name);
    auto ws = wb.sheet_by_index(sheet_index);
    MRange mrng(range_ref);
    size_t found = xlsx_name.find_last_of("/\\");
    outws.cell(last_col + 1, mrng.begin_row).value(xlsx_name.substr(found + 1));
    for (int irow = mrng.begin_row; irow <= mrng.end_row; ++irow) {
        for (int icol = mrng.begin_col; icol <= mrng.end_col; ++icol) {
            copy_value(outws, last_col + 1, irow + 1, ws, icol, irow);
        }
    }
    last_col++;
}

int CollectXLSX::get_1index_collected() {
    for (int i = 0; i < (int)collected.size(); i++) {
        if (collected[i]) return i;
    }
    return -1;
}

void CollectXLSX::add_1xlsx_column() {
    xlnt::workbook wb;
    int wsi = get_1index_collected();
    if (wsi >= 0) {
        wb.load(input_xlsx[wsi]);
        auto ws = wb.sheet_by_index(sheet_index);
        MRange mrng(range_1ref);
        for (int irow = mrng.begin_row; irow <= mrng.end_row; ++irow) {
            for (int icol = mrng.begin_col; icol <= mrng.end_col; ++icol) {
                copy_value(outws, 1, irow + 1, ws, icol, irow);
            }
        }
    }
}

void CollectXLSX::create_input_xlsx() {
    string xlsx_ext = ".xlsx";
    for (const auto &file : fs::directory_iterator(input_dir)) {
        if ((file.path().extension() == xlsx_ext) && (output_name != file.path().filename())) {
            input_xlsx.push_back(file.path().string());
            collected.push_back(true);
        }
    }
    created_input = true;
}

void CollectXLSX::collect() {
    if (!created_input) {
        create_input_xlsx();
    }
    add_1xlsx_column();
    for (size_t i = 0; i < collected.size(); ++i) {
        printf("%s\n", input_xlsx[i].c_str());
        if (collected[i]) {
            add_xlsx_column(input_xlsx[i]);
        }
    }
    outwb.save(input_dir + "/" + output_name);
}

void CollectXLSX::clear() {
    input_xlsx.clear();
    collected.clear();
    outwb.clear();
    outwb = xlnt::workbook();
    outws = outwb.active_sheet();
    outws.title("Data");
    created_input = false;
    last_col = 1;
}
