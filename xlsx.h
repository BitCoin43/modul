#include <OpenXLSX.hpp>
#include <vector>


using namespace OpenXLSX;
using std::vector;

class document {
public:
    document(const std::string docName) {
        name = docName;
        if (name.size() > 4) {
            std::string type = ".xlsx";
            short predp = name.size() - 5;
            bool typed = true;
            for (char i = 0; i < 5; i++) {
                typed = (type[i] == name[predp + i]) && typed;
            }
            if (!typed) {
                name = name + ".xlsx";
            }
        }
        else {
            name = name + ".xlsx";
        }
        doc.create(name);
    }
    ~document() {}
    void save() {
        doc.workbook().deleteSheet("Sheet1");
        doc.save();
        doc.close();
    }
    std::string getName() {
        return name;
    }
public:
    class Sheet {
    public:
        Sheet(document* doc, std::string name) {
            this->doc = doc;
            this->name = name;
        }
        void put_str(std::string x, std::string str) {
            doc->doc.workbook().worksheet(name).cell(x).value() = str;
        }
        void put_num(std::string x, long long num) {
            doc->doc.workbook().worksheet(name).cell(x).value() = num;
        }
        void put_dbl(std::string x, double dbl) {
            doc->doc.workbook().worksheet(name).cell(x).value() = dbl;
        }
    private:
        document* doc;
        std::string name;
    };
    Sheet* new_sheet(std::string name) {
        doc.workbook().addWorksheet(name);
        Sheet* lp = new Sheet(this, name);
        sheetsn.push_back(name);
        sheetsp.push_back(lp);
        return lp;
    }
    Sheet* get_sheet(std::string name) {
        Sheet* lp = nullptr;
        for (int i = 0; i < sheetsn.size(); i++) {
            if (sheetsn[i] == name) {
                lp = sheetsp[i];
                break;
            }
        }
        return lp;
    }
private:
    std::string name;
    XLDocument doc;
    vector<std::string> sheetsn;
    vector<Sheet*> sheetsp;
};