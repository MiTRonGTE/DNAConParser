import xml.etree.ElementTree as ET
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression


def read_xml_file(file_name):
    tree = ET.parse(f'./input/{file_name}.xml')
    root = tree.getroot()
    x = "{urn:schemas-microsoft-com:office:spreadsheet}"
    path = f"{x}Worksheet/{x}Table/{x}Row"
    data = {}

    for i in root.findall(path):
        if 55 < int(i.attrib[f"{x}Index"]) < 63:
            for index, value in enumerate(i):
                if index == 0:
                    row = value[0].text
                    data[row] = []
                else:
                    data[row].append(int(value[0].text))
    return data


def plot_and_output_data(data):
    x = [100, 50, 25, 12.5, 6.25]
    y = [((data["A"][i] - data["A"][5]) + (data["A"][i + 6] - data["A"][11])) / 2 for i in range(0, 5)]
    plt.scatter(x, y)
    plt.plot(x, y)
    plt.ticklabel_format(style="plain")
    plt.ylabel("Absorbance")
    plt.xlabel("DNA concentration (ug)")
    plt.show()

    x_reg = [[i] for i in x]
    y_reg = y

    regressor = LinearRegression()
    regressor.fit(x_reg, y_reg)
    intercept = float(regressor.intercept_)
    coef = float(regressor.coef_)

    con = data.copy()
    con.pop("A")
    for key in con.keys():
        con[key] = [str(round(((i - intercept) / coef), 2)) for i in data[key]]

    return con


def output_data(con, file_name):
    output = ""
    for key, value in con.items():
        output += f"{key}\t" + "\t".join(value)
        output += "\n"

    with open(f"./output/{file_name}.tsv", "w") as ff:
        ff.write(output.replace(".", ","))


if __name__ == "__main__":
    import os
    files = []
    for file in os.listdir("./input"):
        if file.endswith(".xml"):
            if file.replace("xml", "tsv") in os.listdir("./output"):
                continue
            files.append(file.replace('.xml', ''))

    for file in files:
        output_data(plot_and_output_data(read_xml_file(file)), file)
