import os, sys
import copy
import xlrd, xlwt, xlutils
from tempfile import TemporaryFile

shop_name_list = [
"BEIJING CHINA WORLD",
"BEIJING SHIN KONG PLACE 1F",
"BEIJING SHIN KONG PLACE 4F",
"CHANGCHUN",
"CHENGDU IFS",
"GUANGZHOU TAIKOO HUI",
"HANGZHOU",
"HANGZHOU TOWER",
"HARBIN",
"NANJING",
"SHANGHAI AVENUE",
"SHANGHAI PENINSULA",
"SHENYANG",
"WUHAN"
]


class LLG_order_summary_item(object):

    def __init__(self):
        self.Business = None
        self.Category = None
        self.Style = None
        self.Fabric = None
        self.ItemCode = None
        self.Item = None
        self.Composition = None
        self.RetailPrice = None
        self.Currency = None
        self.merchandise_list = []

    #Set attributes
    def set_Business(self, business):
        self.Business = business

    def set_Category(self, category):
        self.Category = category

    def set_Style(self, style):
        self.Style = style

    def set_Fabric(self, fabric):
        self.Fabric = fabric

    def set_ItemCode(self, itemcode):
        self.ItemCode = itemcode

    def set_Item(self, item):
        self.Item = item

    def set_Composition(self, composition):
        self.Composition = composition

    def set_RetailPrice(self, retailprice):
        self.RetailPrice = retailprice

    def set_Currency(self, currency):
        self.Currency = currency


    #Get attributes
    def get_Business(self):
        return self.Business

    def get_Category(self):
        return self.Category

    def get_Style(self):
        return self.Style

    def get_Fabric(self):
        return self.Fabric

    def get_ItemCode(self):
        return self.ItemCode

    def get_Item(self):
        return self.Item

    def get_Composition(self):
        return self.Composition

    def get_RetailPrice(self):
        return self.RetailPrice

    def get_Currency(self):
        return self.Currency

    def get_merchandise(self):
        return self.merchandise_list

    #Merchandise Manage
    def add_merchandise(self, merchan):
        self.merchandise_list.append(merchan)

    def show_all(self):
        print "\nBusiness:", self.Business
        print "Category:", self.Category
        print "Style:", self.Style
        print "Fabric:", self.Fabric
        print "Item Code:", self.ItemCode
        print "Item:", self.Item
        print "Composition:", self.Composition
        print "Retail Price:", self.RetailPrice
        print "Currency:", self.Currency
        for item in self.merchandise_list:
            item.show_merchan()


class LLG_orders_overview_item(object):

    def __init__(self, *param):
        self.Business = param[0]
        self.Category = param[1]
        self.ItemCode = param[2]
        self.ItemDescription = param[3]
        self.Col = param[4]
        self.Total = param[5]
        self.StoreChina = param[6]
        self.China = param[7]
        self.store_sells = {
        "BEIJING SHIN KONG PLACE 1F": param[8],
        "BEIJING CHINA WORLD": param[9],
        "GUANGZHOU TAIKOO HUI": param[10],
        "SHANGHAI AVENUE": param[11],
        "CHENGDU IFS": param[12],
        "SHANGHAI PENINSULA": param[13],
        "NANJING": param[14],
        "BEIJING SHIN KONG PLACE 4F": param[15],
        "CHANGCHUN": param[16],
        "HANGZHOU TOWER": param[17],
        "HARBIN": param[18],
        "WUHAN": param[19],
        "SHENYANG": param[20]
        }

    def get_Business(self):
        return self.Business

    def get_Category(self):
        return self.Category

    def get_ItemCode(self):
        return self.ItemCode

    def get_ItemDescription(self):
        return self.ItemDescription

    def get_Col(self):
        return self.Col

    def get_Total(self):
        return self.Total

    def get_StoreChina(self):
        return self.StoreChina

    def get_China(self):
        return self.China

    def get_store_sells(self):
        return self.store_sells

    def show_all(self):
        print "\nBusiness:", self.Business
        print "Category:", self.Category
        print "Item Code:", self.ItemCode
        print "Item Description:", self.ItemDescription
        print "Col:", self.Col
        print "Total:", self.Total
        print "Store China:", self.StoreChina   
        print "China:", self.China
        print "Store sells:"
        print self.store_sells


class merchandise_item(object):

    #0 is Net Qty, 1 is Shipped
    NR_static = [0, 0]

    def __init__(self, color_code, color, delivery_from, delivery_to):
        self.color_code = color_code
        self.color = color
        self.delivery_from = delivery_from
        self.delivery_to = delivery_to
        self.shop_list = []
        self.NR = [0, 0]

    def get_color_code(self):
        return self.color_code

    def get_color(self):
        return self.color

    def get_delivery_from(self):
        return self.delivery_from

    def get_delivery_to(self):
        return self.delivery_to

    def get_shop_list(self):
        return self.shop_list

    def get_NR(self):
        return self.NR

    def add_shop(self, shop_name, NetQty, Shipped):
        self.shop_list.append((shop_name, NetQty, Shipped))
        self.NR[0] += int(NetQty)
        self.NR[1] += int(Shipped)
        self.NR_static[0] += int(NetQty)
        self.NR_static[1] += int(Shipped)

    def show_merchan(self):
        print "Color Code:", self.color_code
        print "Color:", self.color
        print "Delivery From:", self.delivery_from
        print "Delivery To:", self.delivery_to
        for item in self.shop_list:
            print "\tshop name:", item[0]
            print "\tNetQty:", item[1]
            print "\tShipped:", item[2]


def read_LLG_orders_overview(LLG_orders_overview):

    current_path = sys.path[0]
    LLG_order_overview_path = os.path.join(current_path, LLG_orders_overview)
    LLG_order_overview_book = xlrd.open_workbook(LLG_order_overview_path)
    LLG_order_overview_table = LLG_order_overview_book.sheets()[0]

    data = []
    for rx in range(LLG_order_overview_table.nrows):
        line = []
        for ry in range(LLG_order_overview_table.ncols):
            cell = ''
            try:
                cell = LLG_order_overview_table.cell_value(rx, ry)
            except:
                pass
            line.append(cell)
        if len(line) > 0:
            data.append(line)

    total_LLG_orders_overview_list = []

    line_num = 0

    while data[line_num][3] != "TOTAL":
        line_num += 1

    while line_num < len(data):
        line_num += 1
        temp_business = data[line_num][3]
        while line_num < len(data):
            line_num += 1
            if data[line_num][4] != "":
                temp_category = data[line_num][4]
                while line_num < len(data):
                    line_num += 1
                    if data[line_num][5] != "" and data[line_num][7] != "":
                        temp_itemcode = data[line_num][5]
                        temp_itemdescription = data[line_num][7]
                        while line_num < len(data):
                            line_num += 1
                            if line_num >= len(data):
                                break
                            if data[line_num][8] != "":
                                temp_item = LLG_orders_overview_item(\
                                    temp_business, temp_category, temp_itemcode, \
                                    temp_itemdescription, *data[line_num][8:25])
                                total_LLG_orders_overview_list.append(temp_item)
                                #temp_item.show_all()
                            else:
                                line_num -= 1
                                break
                    else:
                        line_num -= 1
                        break
            else:
                line_num -= 1
                break

    return total_LLG_orders_overview_list


def read_LLG_order_summary(LLG_order_summary):

    current_path = sys.path[0]
    LLG_order_summary_path = os.path.join(current_path, LLG_order_summary)
    LLG_order_summary_book = xlrd.open_workbook(LLG_order_summary_path)
    LLG_order_summary_table = LLG_order_summary_book.sheets()[0]

    data = []
    for rx in range(LLG_order_summary_table.nrows):
        line = []
        for ry in range(LLG_order_summary_table.ncols):
            cell = ''
            try:
                cell = LLG_order_summary_table.cell_value(rx, ry)
            except:
                pass
            if cell != '':
                line.append(cell)
        if len(line) > 0:
            data.append(line)

    total_LLG_order_summary_list = []

    temp_LLG_item = 0

    line_num = 0

    while line_num < len(data):

        if data[line_num][0] == "Business":
            if temp_LLG_item == 0:
                temp_LLG_item = LLG_order_summary_item()
                temp_LLG_item.set_Business(data[line_num][1])
            else:
                print "temp_LLG_item error."
                sys.exit()

        if data[line_num][0] == "Category":
            temp_LLG_item.set_Category(data[line_num][1])

        if data[line_num][0] == "Style":
            temp_LLG_item.set_Style(data[line_num][1])

        if data[line_num][0] == "Fabric":
            temp_LLG_item.set_Fabric(data[line_num][1])

        if data[line_num][0] == "Item Code":
            temp_LLG_item.set_ItemCode(data[line_num][1])

        if data[line_num][0] == "Item":
            temp_LLG_item.set_Item(data[line_num][1])

        if data[line_num][0] == "Composition":
            temp_LLG_item.set_Composition(data[line_num][1])

        if data[line_num][0] == "Currency":
            temp_LLG_item.set_Currency(data[line_num][1])

        if data[line_num][0] == "Retail Price":
            temp_LLG_item.set_RetailPrice(data[line_num][1])

        if data[line_num][0] == "UM":

            temp_mer = 0

            line_num += 1
            while line_num < len(data):

                if len(data[line_num]) == 8:

                    if temp_mer != 0:
                        #temp_mer.show_merchan()
                        temp_LLG_item.add_merchandise(temp_mer)
                        
                    temp_mer = merchandise_item(*data[line_num][1:5])
                    temp_mer.add_shop(*data[line_num][5:8])

                if len(data[line_num]) == 7:

                    #temp_mer.show_merchan()
                    temp_LLG_item.add_merchandise(temp_mer)

                    temp_mer = merchandise_item(*data[line_num][0:4])
                    temp_mer.add_shop(*data[line_num][4:7])
                
                if len(data[line_num]) == 3:
                    if data[line_num][0] != "NR":
                        temp_mer.add_shop(*data[line_num])
                    else:
                        #temp_mer.show_merchan()
                        temp_LLG_item.add_merchandise(temp_mer)
                        total_LLG_order_summary_list.append(temp_LLG_item)
                        temp_LLG_item = 0
                        break

                line_num += 1

        line_num += 1

    return total_LLG_order_summary_list


class LRTW_order_summary_item(object):

    def __init__(self):
        self.Business = None
        self.Category = None
        self.Style = None
        self.Fabric = None
        self.ItemCode = None
        self.Item = None
        self.Composition = None
        self.RetailPrice = None
        self.Currency = None
        self.size_type = None
        self.merchandise_list = []

    #Set attributes
    def set_Business(self, business):
        self.Business = business

    def set_Category(self, category):
        self.Category = category

    def set_Style(self, style):
        self.Style = style

    def set_Fabric(self, fabric):
        self.Fabric = fabric

    def set_ItemCode(self, itemcode):
        self.ItemCode = itemcode

    def set_Item(self, item):
        self.Item = item

    def set_Composition(self, composition):
        self.Composition = composition

    def set_RetailPrice(self, retailprice):
        self.RetailPrice = retailprice

    def set_Currency(self, currency):
        self.Currency = currency

    def set_size_type(self, size_type):
        self.size_type = size_type


    #Get attributes
    def get_Business(self):
        return self.Business

    def get_Category(self):
        return self.Category

    def get_Style(self):
        return self.Style

    def get_Fabric(self):
        return self.Fabric

    def get_ItemCode(self):
        return self.ItemCode

    def get_Item(self):
        return self.Item

    def get_Composition(self):
        return self.Composition

    def get_RetailPrice(self):
        return self.RetailPrice

    def get_Currency(self):
        return self.Currency

    def get_merchandise(self):
        return self.merchandise_list

    def get_size_type(self):
        return self.size_type

    #Merchandise Manage
    def add_merchandise(self, merchan):
        self.merchandise_list.append(merchan)

    def show_all(self):
        print "\nBusiness:", self.Business
        print "Category:", self.Category
        print "Style:", self.Style
        print "Fabric:", self.Fabric
        print "Item Code:", self.ItemCode
        print "Item:", self.Item
        print "Composition:", self.Composition
        print "Retail Price:", self.RetailPrice
        print "Currency:", self.Currency
        for item in self.merchandise_list:
            item.show_merchan()

    def find_sell(self, color_code, shop_name):
        for merchan in self.merchandise_list:
            if merchan.get_color_code() == color_code:
                for shop in merchan.get_shop_list():
                    if shop[0] == shop_name:
                        if self.size_type == "A":
                            return [shop[1]]
                        if self.size_type == "B":
                            return shop[2:7]
                        if self.size_type == "C":
                            return shop[7:]
        return None


class LRTW_orders_overview_item(object):

    def __init__(self, *param):
        self.Business = param[0]
        self.Category = param[1]
        self.ItemCode = param[2]
        self.ItemDescription = param[3]
        self.Col = param[4]
        self.Total = param[5]
        self.StoreChina = param[6]
        self.China = param[7]
        self.store_sells = {
        "BEIJING SHIN KONG PLACE 1F": param[8],
        "BEIJING CHINA WORLD": param[9],
        "GUANGZHOU TAIKOO HUI": param[10],
        "SHANGHAI AVENUE": param[11],
        "CHENGDU IFS": param[12],
        "SHANGHAI PENINSULA": param[13],
        "NANJING": param[14],
        "BEIJING SHIN KONG PLACE 4F": param[15],
        "CHANGCHUN": param[16],
        "HANGZHOU TOWER": param[17],
        "HARBIN": param[18],
        "WUHAN": param[19],
        "SHENYANG": param[20]
        }

    def get_Business(self):
        return self.Business

    def get_Category(self):
        return self.Category

    def get_ItemCode(self):
        return self.ItemCode

    def get_ItemDescription(self):
        return self.ItemDescription

    def get_Col(self):
        return self.Col

    def get_Total(self):
        return self.Total

    def get_StoreChina(self):
        return self.StoreChina

    def get_China(self):
        return self.China

    def get_store_sells(self):
        return self.store_sells

    def show_all(self):
        print "\nBusiness:", self.Business
        print "Category:", self.Category
        print "Item Code:", self.ItemCode
        print "Item Description:", self.ItemDescription
        print "Col:", self.Col
        print "Total:", self.Total
        print "Store China:", self.StoreChina   
        print "China:", self.China
        print "Store sells:"
        print self.store_sells


class clothes_merchandise_item(object):

    def __init__(self, color_code, color, size_type):
        self.color_code = color_code
        self.color = color
        self.size_type = size_type
        self.shop_list = []

    def get_color_code(self):
        return self.color_code

    def get_color(self):
        return self.color

    def get_size_type(self):
        return self.size_type

    def get_shop_list(self):
        return self.shop_list


    def add_shop(self, reference_line, line, size_type):


        #    0     ,  1  ,  2  ,  3 ,  4 ,  5 ,  6  ,  7  ,  8  ,  9  ,  10 ,  11 ,  12 ,  13
        # shop_name, A_NR, B_XS, B_S, B_M, B_L, B_XL, C_36, C_38, C_40, C_42, C_44, C_46, C_48
        param_list = ["" for x in range(14)]

        param_list[0] = line[2]

        if size_type == "A":
            param_list[1] = line[3]

        if size_type == "B":
            for index in range(3, len(reference_line)):
                if reference_line[index] == "XS":
                    param_list[2] = line[index]
                if reference_line[index] == "S":
                    param_list[3] = line[index]
                if reference_line[index] == "M":
                    param_list[4] = line[index]
                if reference_line[index] == "L":
                    param_list[5] = line[index]
                if reference_line[index] == "XL":
                    param_list[6] = line[index]

        if size_type == "C":
            for index in range(3, len(reference_line)):
                if reference_line[index] == "36":
                    param_list[7] = line[index]
                if reference_line[index] == "38":
                    param_list[8] = line[index]
                if reference_line[index] == "40":
                    param_list[9] = line[index]
                if reference_line[index] == "42":
                    param_list[10] = line[index]
                if reference_line[index] == "44":
                    param_list[11] = line[index]
                if reference_line[index] == "46":
                    param_list[12] = line[index]
                if reference_line[index] == "48":
                    param_list[13] = line[index]

        self.shop_list.append(param_list)

    def show_merchan(self):
        print "Color Code:", self.color_code
        print "Color:", self.color
        for item in self.shop_list:
            print "\tshop:", item


def read_LRTW_orders_overview(LRTW_orders_overview):

    current_path = sys.path[0]
    LRTW_order_overview_path = os.path.join(current_path, LRTW_orders_overview)
    LRTW_order_overview_book = xlrd.open_workbook(LRTW_order_overview_path)
    LRTW_order_overview_table = LRTW_order_overview_book.sheets()[0]

    data = []
    for rx in range(LRTW_order_overview_table.nrows):
        line = []
        for ry in range(LRTW_order_overview_table.ncols):
            cell = ''
            try:
                cell = LRTW_order_overview_table.cell_value(rx, ry)
            except:
                pass
            line.append(cell)
        if len(line) > 0:
            data.append(line)

    total_LRTW_orders_overview_list = []

    line_num = 0

    while data[line_num][3] != "TOTAL":
        line_num += 1

    while line_num < len(data):
        line_num += 1
        temp_business = data[line_num][3]
        while line_num < len(data):
            line_num += 1
            if data[line_num][4] != "":
                temp_category = data[line_num][4]
                while line_num < len(data):
                    line_num += 1
                    if data[line_num][5] != "" and data[line_num][7] != "":
                        temp_itemcode = data[line_num][5]
                        temp_itemdescription = data[line_num][7]
                        while line_num < len(data):
                            line_num += 1
                            if line_num >= len(data):
                                break
                            if data[line_num][8] != "":
                                temp_item = LRTW_orders_overview_item(\
                                    temp_business, temp_category, temp_itemcode, \
                                    temp_itemdescription, *data[line_num][8:25])
                                total_LRTW_orders_overview_list.append(temp_item)
                                #temp_item.show_all()
                            else:
                                line_num -= 1
                                break
                    else:
                        line_num -= 1
                        break
            else:
                line_num -= 1
                break

    return total_LRTW_orders_overview_list


def read_LRTW_order_summary(LRTW_order_summary):

    current_path = sys.path[0]
    LRTW_order_summary_path = os.path.join(current_path, LRTW_order_summary)
    LRTW_order_summary_book = xlrd.open_workbook(LRTW_order_summary_path)
    LRTW_order_summary_table = LRTW_order_summary_book.sheets()[0]

    data = []
    for rx in range(LRTW_order_summary_table.nrows):
        line = []
        for ry in range(LRTW_order_summary_table.ncols):
            cell = ''
            try:
                cell = LRTW_order_summary_table.cell_value(rx, ry)
            except:
                pass
            line.append(cell)
        if len(line) > 0:
            data.append(line)

    total_LRTW_order_summary_list = []

    temp_LRTW_item = 0

    line_num = 0
    
    while line_num < len(data):

        if data[line_num][0] == "Business":
            if temp_LRTW_item == 0:
                temp_LRTW_item = LRTW_order_summary_item()
                temp_LRTW_item.set_Business(data[line_num][2])
            else:
                print "temp_LRTW_item error."
                sys.exit()

        if data[line_num][0] == "Category":
            temp_LRTW_item.set_Category(data[line_num][2])

        if data[line_num][0] == "Style":
            temp_LRTW_item.set_Style(data[line_num][2])

        if data[line_num][0] == "Fabric":
            temp_LRTW_item.set_Fabric(data[line_num][2])

        if data[line_num][0] == "Item Code":
            temp_LRTW_item.set_ItemCode(data[line_num][2])

        if data[line_num][0] == "Item":
            temp_LRTW_item.set_Item(data[line_num][2])

        if data[line_num][0] == "Composition":
            temp_LRTW_item.set_Composition(data[line_num][2])

        if data[line_num][0] == "Currency":
            temp_LRTW_item.set_Currency(data[line_num][2])

        if data[line_num][0] == "Retail Price":
            temp_LRTW_item.set_RetailPrice(data[line_num][2])

        if data[line_num][0] == "Color Code":

            reference_line = data[line_num]

            A_set = set(["NR"])
            B_set = set(["XS", "S", "M", "L", "XL"])
            C_set = set(["36", "38", "40", "42", "44", "46", "48"])

            size_type = 0

            if reference_line[4] in A_set:
                size_type = "A"
            if reference_line[4] in B_set:
                size_type = "B"
            if reference_line[4] in C_set:
                size_type = "C"

            temp_LRTW_item.set_size_type(size_type)

            temp_mer = 0

            line_num += 1

            while line_num < len(data):

                new_reference_line = copy.deepcopy(reference_line)

                index = 0
                while index < len(new_reference_line):
                    if new_reference_line[index] == "" and data[line_num][index] == "":
                        del new_reference_line[index]
                        del data[line_num][index]
                    else:
                        index += 1

                #print "\nRef:", new_reference_line
                #print "Raw:", data[line_num]
                #print "\n"

                if data[line_num][0] != "" and data[line_num][1] != "":

                    if temp_mer != 0:
                        #temp_mer.show_merchan()
                        temp_LRTW_item.add_merchandise(temp_mer)
                    

                    temp_mer = clothes_merchandise_item(data[line_num][0], data[line_num][1], size_type)

                    temp_mer.add_shop(new_reference_line, data[line_num], size_type)

                if data[line_num][2] == "Total Net Qty":
                    line_num += 1
                    continue

                if data[line_num][0] == "" and data[line_num][1] == "" \
                    and data[line_num][2].find("Total Net Qty of") < 0:
                    temp_mer.add_shop(new_reference_line, data[line_num], size_type)
                
                if data[line_num][2].find("Total Net Qty of") >= 0:
                    temp_LRTW_item.add_merchandise(temp_mer)
                    total_LRTW_order_summary_list.append(temp_LRTW_item)
                    temp_LRTW_item = 0
                    break

                line_num += 1

        line_num += 1
    
    return total_LRTW_order_summary_list


def load_format(sample_filename):

    current_path = sys.path[0]
    sample_path = os.path.join(current_path, sample_filename)

    sample_book = xlrd.open_workbook(sample_path)
    sample_table = sample_book.sheets()[0]

    output_book = xlwt.Workbook()

    sheet1 = output_book.add_sheet('Sheet 1')

    for xrow in range(sample_table.nrows - 1):
        for xcol in range(sample_table.ncols):
            sheet1.write(xrow,xcol,sample_table.cell_value(xrow, xcol))

    return sheet1, output_book


def fill_size_data(sheet, size_type, line_num, start, data):

    if data == None:
        return

    total_num = 0

    if size_type == "A":
        sheet.write(line_num, start, data[0])
        total_num = data[0]
        
    if size_type == "B":
        for index in range(5):
            sheet.write(line_num, start + index, data[index])
        total_num = sum([int(x) for x in data if x != ""])

    if size_type == "C":
        for index in range(7):
            sheet.write(line_num, start + index, data[index])
        total_num = sum([int(x) for x in data if x != ""])

    sheet.write(line_num, start + 7, total_num)


def main():

    #if len(sys.argv) < 4:
        #print "Please include the name of Excel files in the command line."
        #sys.exit()

    LLG_order_summary = "SIN_61_Order_Summary_LLG.xls"
    LLG_orders_overview = "SIN_02_Orders_Overview_LLG.xls"
    LRTW_order_summary = "SIN_63_Order_Summary_by_Size_LRTW.xls"
    LRTW_orders_overview = "SIN_02_Orders_Overview_LRTW.xls"
    sample_filename = "SS16 LINELIST LRTW SAMPLE.xlsx"


    LLG_order_summary_list = read_LLG_order_summary(LLG_order_summary)

    LLG_orders_overview_list = read_LLG_orders_overview(LLG_orders_overview)

    LRTW_order_summary_list = read_LRTW_order_summary(LRTW_order_summary)

    LRTW_orders_overview_list = read_LRTW_orders_overview(LRTW_orders_overview)

    sheet1, output_book = load_format(sample_filename)

    line_num = 4

    for summary_item in LLG_order_summary_list:

        #Find the pair
        for order_item in summary_item.get_merchandise():
            for overview_item in LLG_orders_overview_list:
                if summary_item.get_ItemCode() == overview_item.get_ItemCode() and \
                    order_item.get_color_code() == overview_item.get_Col():

                    #COLLECTION
                    sheet1.write(line_num, 0, "")

                    #BUSINESS
                    sheet1.write(line_num, 1, summary_item.get_Business())

                    #CAT
                    sheet1.write(line_num, 2, "")

                    #CATEGORY
                    sheet1.write(line_num, 3, summary_item.get_Category())

                    #SKETCHES
                    sheet1.write(line_num, 4, "")

                    #STYLE
                    sheet1.write(line_num, 5, summary_item.get_ItemCode())

                    #STYLE DESCRIPTION
                    sheet1.write(line_num, 6, overview_item.get_ItemDescription())

                    #COLOR
                    sheet1.write(line_num, 7, overview_item.get_Col())

                    #COLOR DESCRIPTION
                    sheet1.write(line_num, 8, order_item.get_color())

                    #COLOR CHIPS
                    sheet1.write(line_num, 9, "")

                    #COMPOSITIONS
                    sheet1.write(line_num, 10, summary_item.get_Composition())

                    #Retail Price RMB
                    sheet1.write(line_num, 11, summary_item.get_RetailPrice())

                    #OPTION
                    sheet1.write(line_num, 12, "A")

                    #PRC TOTAL
                    sheet1.write(line_num, 13, overview_item.get_China())

                    #PRC TOTAL
                    sheet1.write(line_num, 20, overview_item.get_China())

                    #Shops
                    count = 21
                    for name in shop_name_list:
                        if name != "HANGZHOU":
                            sheet1.write(line_num, count, overview_item.get_store_sells()[name])
                            count += 8

                    line_num += 1
                    break
    
    for summary_item in LRTW_order_summary_list:

        #Find the pair
        for order_item in summary_item.get_merchandise():
            for overview_item in LRTW_orders_overview_list:
                if summary_item.get_ItemCode() == overview_item.get_ItemCode() and \
                    order_item.get_color_code() == overview_item.get_Col():

                    #COLLECTION
                    sheet1.write(line_num, 0, "")

                    #BUSINESS
                    sheet1.write(line_num, 1, summary_item.get_Business())

                    #CAT
                    sheet1.write(line_num, 2, "")

                    #CATEGORY
                    sheet1.write(line_num, 3, summary_item.get_Category())

                    #SKETCHES
                    sheet1.write(line_num, 4, "")

                    #STYLE
                    sheet1.write(line_num, 5, summary_item.get_ItemCode())

                    #STYLE DESCRIPTION
                    sheet1.write(line_num, 6, overview_item.get_ItemDescription())

                    #COLOR
                    sheet1.write(line_num, 7, overview_item.get_Col())

                    #COLOR DESCRIPTION
                    sheet1.write(line_num, 8, order_item.get_color())

                    #COLOR CHIPS
                    sheet1.write(line_num, 9, "")

                    #COMPOSITIONS
                    sheet1.write(line_num, 10, summary_item.get_Composition())

                    #Retail Price RMB
                    sheet1.write(line_num, 11, summary_item.get_RetailPrice())

                    #OPTION
                    sheet1.write(line_num, 12, summary_item.get_size_type())

                    #PRC TOTAL
                    total_list = []

                    if summary_item.get_size_type() == "A":
                        total = 0
                        for shop in shop_name_list:
                            temp = summary_item.find_sell(order_item.get_color_code(), shop)
                            if temp != None:
                                total += int(temp[0])
                        total_list.append(total)

                    if summary_item.get_size_type() == "B":
                        total_list = [0 for x in range(5)]
                        for shop in shop_name_list:
                            temp = summary_item.find_sell(order_item.get_color_code(), shop)
                            if temp != None:
                                for index in range(5):
                                    if temp[index] != "":
                                        total_list[index] += int(temp[index])

                    if summary_item.get_size_type() == "C":
                        total_list = [0 for x in range(7)]
                        for shop in shop_name_list:
                            temp = summary_item.find_sell(order_item.get_color_code(), shop)
                            if temp != None:
                                for index in range(7):
                                    if temp[index] != "":
                                        total_list[index] += int(temp[index])

                    for index in range(len(total_list)):
                        if total_list[index] == 0:
                            total_list[index] = ""

                    fill_size_data(sheet1, summary_item.get_size_type(), line_num, 13, total_list)

                    #Shops
                    count = 21
                    for name in shop_name_list:
                        fill_size_data(sheet1, summary_item.get_size_type(), line_num, count,\
                        summary_item.find_sell(order_item.get_color_code(), name))
                        count += 8

                    line_num += 1
                    break


    output_path = os.path.join(sys.path[0], "Output.xls")
    output_book.save(output_path)
    output_book.save(TemporaryFile())



    '''
    sheet1.write(0,0,'A1')
    sheet1.write(0,1,'B1')
    row1 = sheet1.row(1)
    row1.write(0,'A2')
    row1.write(1,'B2')
    sheet1.col(0).width = 10000
    sheet2 = book_2.get_sheet(1)
    sheet2.row(0).write(0,'Sheet 2 A1')
    sheet2.row(0).write(1,'Sheet 2 B1')
    sheet2.flush_row_data()
    sheet2.write(1,0,'Sheet 2 A3')
    sheet2.col(0).width = 5000
    sheet2.col(0).hidden = False
    '''


if __name__ == '__main__':
    main()