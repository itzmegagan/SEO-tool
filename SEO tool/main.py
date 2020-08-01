from urllib.request import urlopen
import re
from bs4 import BeautifulSoup
import sqlite3
import xlsxwriter

workbook = xlsxwriter.Workbook("seo_tool.xlsx")


def read_database(cnt):
    conn = sqlite3.connect("seo_project.db")
    c = conn.cursor()
    c.execute("SELECT * FROM 'seo_project' LIMIT {start},{end}".format(start=cnt, end=cnt + 10))
    all_rows = c.fetchall()
    return all_rows


def write_column(cnt, url):
    toplist = read_database(cnt)

    sheetname = url[-8:]
    sheetname = sheetname.replace("/", "")
    sheetname = sheetname.replace(".", "")
    sheetname = sheetname.replace("[", "")
    sheetname = sheetname.replace("]", "")
    sheetname = sheetname.replace("(", "")
    sheetname = sheetname.replace(")", "")
    worksheet = workbook.add_worksheet(sheetname)

    bold = workbook.add_format({'bold': True})

    worksheet.write("A1", "WORD", bold)
    worksheet.write("C1", "FREQUENCY", bold)
    worksheet.write("E1", "DENSITY", bold)

    row = 1
    col = 0

    for item, freq, dens in (toplist):
        worksheet.write(row, col, item)
        worksheet.write(row, col + 2, freq)
        worksheet.write(row, col + 4, dens)
        row = row + 1

    chart1 = workbook.add_chart({'type': 'pie'})

    chart1.add_series({
        'name': 'Pie frequency data',
        'categories': '=' + sheetname + '!' + '$A$1',
        'values': '=' + sheetname + '!' + '$C$2:$C$7',
        'points': [
            {'fill': {'color': '#5ABA10'}},
            {'fill': {'color': '#FE110E'}},
            {'fill': {'color': '#CA5C05'}},
            {'fill': {'color': '#FF9900'}},
            {'fill': {'color': '#800080'}}]})

    chart1.set_title({'name': 'Pie Chart with top six frequent words'})

    worksheet.insert_chart('G2', chart1)

    chart2 = workbook.add_chart({'type': 'line'})

    x_vals = sheetname + '!' + '$A$2:$A$11'
    freq_vals = sheetname + '!' + '$C$2:$C$11'
    dens_vals = sheetname + '!' + '$E$2:$E$11'

    chart2.add_series({
        'name': '=' + sheetname + '!' + '$BC$1',
        'categories': x_vals,
        'values': freq_vals,
        'line': {'color': 'green'},
    })

    chart2.add_series({
        'name': '=' + sheetname + '!' + '$E$1',
        'categories': x_vals,
        'values': dens_vals,
        'line': {'color': 'orange'},
        'y2_axis': True,
    })

    chart2.set_y_axis({'name': 'Frequency'})
    chart2.set_y2_axis({'name': 'Density'})
    chart2.set_size({'width': 600, 'height': 300})

    worksheet.insert_chart('O2', chart2)


cnt = 0

whandle = open("C:\\Users\\DrBekal\\projectseo\\websiteurl.txt")
for url in whandle:
    pattern = re.compile("(https://|http://)[\w]+.[\w]+")
    matchobject = pattern.match(url)

    if not (matchobject):
        print("Invalid url")

    else:

        print("your website is: ", url)
        html = urlopen(url).read()
        soup = BeautifulSoup(html, "html.parser")

        total_text = ""

        for script in soup(["script", "style"]):
            script.extract()

        text = soup.get_text()
        text = text.lower()
        text = text.replace("?", "")
        text = text.replace("^", "")
        text = text.replace("!", "")
        text = text.replace("'", "")
        text = text.replace("+", "")
        text = text.replace("]", "")
        text = text.replace("[", "")
        text = text.replace("}", "")
        text = text.replace("(", "")
        text = text.replace("{", "")
        text = text.replace(")", "")
        text = text.replace("  ", " ")
        text = text.replace('\n', " ")
        text = text.replace(",", "")
        text = text.replace(".", "")
        text = text.replace("-", "")
        text = text.replace(":", "")
        text = text.replace("|", "")
        text = text.replace("-", "")
        text = text.replace("_", "")
        text = text.replace("||", "")

        total_text = total_text + text

        words = total_text.split()

        mydict = {}
        for word in words:
            fhandle = open("C:\\Users\\DrBekal\\projectseo\\ignorewords.txt")
            for line in fhandle:
                ignore = line.split()
                if word not in ignore:
                    if word in mydict:
                        mydict[word] = mydict[word] + 1
                    else:
                        mydict[word] = 1

        count = 0
        for word in words:
            if word not in words:
                count = 1
            else:
                count = count + 1

        for word in mydict:
            frequency = mydict[word]
            density = float(frequency / count * 100.0)


        topwords = sorted(mydict.items(), key=lambda x :x[1], reverse=True)[:10]
        toplist = list(topwords)

        newtoplist = []

        for word, freq in toplist:
            new = []
            new.append(word)
            new.append(freq)
            new.append(float(freq / count) * 100.0)
            newtoplist.append(new)

        conn = sqlite3.connect("seo_project.db")
        cur = conn.cursor()
        cur.execute('''CREATE TABLE IF NOT EXISTS seo_project(word varchar,frequency int,density int)''')
        for item in newtoplist:
            cur.execute('''INSERT INTO seo_project VALUES(?,?,?)''', item)
        conn.commit()
        conn.close()
        write_column(cnt, url)
        cnt = cnt + 10


workbook.close()