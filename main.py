from datetime import datetime, timedelta

### These libraries needs to be uploaded ###
from openpyxl import Workbook
from openpyxl.styles import Font

# ====== ilk olarak kac gunluk veri gerektigini ve en az kaç dk musait zaman istedigimizi cekiyoruz. ======

database_dosya_yolu = "C:\\Users\\furkan.cankaya\\Desktop\\Dosyalar\\ExcelKod\\data.csv"
degiskenlercsv_ortak_dosya_yolu = "C:\\Users\\furkan.cankaya\\Desktop\\Dosyalar\\ExcelKod\\Degiskenler.csv"
utilizasyonraporu_cikti_dosyasi_yolu = "C:\\Users\\furkan.cankaya\\Desktop\\Dosyalar\\ExcelKod\\Utilizasyon.xlsx"

with open(degiskenlercsv_ortak_dosya_yolu, "r", encoding="utf-8") as dosya:
    degisken_satirlari = dosya.readlines()

with open(database_dosya_yolu, "r", encoding="utf-8") as dosya:
    veri_satirlari = dosya.readlines()

day_interval = int(degisken_satirlari[1].split(";")[1])
min_free_time_minutes = int(degisken_satirlari[2].split(";")[1])




i = 0
while i < len(veri_satirlari):
    if veri_satirlari[-(i+1)].split(",")[11].strip() != "":
        lastdate = datetime.strptime(veri_satirlari[-(i+1)].split(",")[11].strip(), "%m/%d/%Y %I:%M:%S %p")
        startdate = lastdate - timedelta(days=day_interval)

        break
    i += 1



titles = veri_satirlari[0].split(",")
duzenlenmis_veri_satirlari = []

# ====== Datayi duzenliyoruz ve sadelestiriyoruz. ======

# Process;Machine;Hostname;Host Identity;Job type;Runtime type/license;State;Priority;Started (absolute);Ended (absolute);Source;Created (absolute)
duzenlenmis_veri_satirlari.append(";".join([titles[1], 
                                            titles[3], 
                                            titles[4], 
                                            titles[6], 
                                            titles[7], 
                                            titles[8], 
                                            titles[9], 
                                            titles[10], 
                                            titles[11], 
                                            titles[12], 
                                            titles[13], 
                                            titles[14][:-1]]))


counter = 0
for satir in veri_satirlari:
    if counter > 0:
        satir_bolunmus = satir.strip().split(",")
        yeni_satir = []
        for eleman in satir_bolunmus:
            yeni_satir.append("")
            yeni_satir.append(eleman)

        try:
            tempdate = datetime.strptime(yeni_satir[23], "%m/%d/%Y %I:%M:%S %p")
            if tempdate > startdate:
                duzenlenmis_veri_satirlari.append(";".join([yeni_satir[3], 
                                                            yeni_satir[7], 
                                                            yeni_satir[9], 
                                                            yeni_satir[13], 
                                                            yeni_satir[15], 
                                                            yeni_satir[17], 
                                                            yeni_satir[19], 
                                                            yeni_satir[21], 
                                                            yeni_satir[23], 
                                                            yeni_satir[25], 
                                                            yeni_satir[27], 
                                                            yeni_satir[29]]))

        except:
            print(yeni_satir[23])
            

    counter += 1


# with open("sorgulanmisveriler.csv", "w", encoding="utf-8-sig") as yeni_dosya:
#     for satir in duzenlenmis_veri_satirlari:
#         yeni_dosya.write(satir + "\n")
print(str(startdate) + " tarihinden " + str(lastdate) + " tarihine kadar olan veriler excel e aktarildi.")


# ============ Verileri robotlara atİyoruz. ============

# ilk olarak bir json veritabanİmİz olacak bu veritabanİnda ROBOT:{name:SeherGida;Status:Success} seklinde veriler tutulacak.
# ilk 1. boyutta robot isimlerinin bulunacagi listeleri olusturmamiz gerekiyor.

robots = []
utilization = {}
counter = 0


# ilk boyutta robot isimlerini topluyoruz
for i in duzenlenmis_veri_satirlari:
    if counter > 0:
        robotname = i.split(";")[3]
        if robotname[11:] not in robots and robotname[11:] != "":
            robots.append(robotname[11:])
    counter += 1

# Robotlar listede gordukleri siraya gore ekleniyordu. Asagidaki fonksiyona gore Robotlarin numaralarina gore atanmasi saglandi. 
def sort_robots(robot_list):
    def robot_key(robot):
        if robot == "ROBOT":
            return 1
        else:
            return int(robot.replace("ROBOT", "")) * 10

    # Robotlari numaralarina gore siraliyoruz
    sorted_list = sorted(robot_list, key=robot_key)
    return sorted_list

robots = sort_robots(robots)

for robot_name in robots:
    utilization[robot_name] = {"processes": [], 
                               "efficiency": [],
                               "available":[]}

# ikinci boyutta robot surelerini dolduruyoruz
counter = 0
for i in duzenlenmis_veri_satirlari:
    if counter > 0:
        data = i.split(";")
        name = data[0]
        robotname = data[3]
        starteddate = data[8]
        endeddate = data[9]
        processresult = data[6]
        processstarter = data[10]

        if robotname[11:] in utilization:
            process_data = {
                "name": name,
                "starteddate": starteddate,
                "endeddate": endeddate,
                "processresult": processresult,
                "processstarter": processstarter
            }
            utilization[robotname[11:]]["processes"].append(process_data)

    counter += 1




def format_time(minutes):
    days = int(minutes // (24 * 60))
    hours = int((minutes % (24 * 60)) // 60)
    mins = int(minutes % 60)
    
    if days > 0:
        return f"{days} gün, {hours} saat, {mins} dakika"
    elif hours > 0:
        return f"{hours} saat, {mins} dakika"
    else:
        return f"{mins} dakika"


def calculate_time_difference_in_minutes(start_date_str, end_date_str):
    start_date = datetime.strptime(start_date_str, "%m/%d/%Y %I:%M:%S %p")
    end_date = datetime.strptime(end_date_str, "%m/%d/%Y %I:%M:%S %p")

    time_difference = end_date - start_date

    time_difference_in_minutes = time_difference.total_seconds() / 60

    return time_difference_in_minutes


bold_font = Font(bold=True)
availabletimes = ""

wb = Workbook()

for robotprocesses in utilization:
    success_counter = 0
    faulted_counter = 0
    pending_counter = 0
    calc_successful = 0
    robot_free_intervals = []
    
    ws = wb.create_sheet(robotprocesses)
    ws.append(["Process","Started (absolute)","Ended (absolute)","Time","Status","Trigger"])
    for cell in ws[1]:
        cell.font = bold_font

    sorted_processes = sorted(utilization[robotprocesses]["processes"], key=lambda x: datetime.strptime(x["starteddate"], "%m/%d/%Y %I:%M:%S %p"))
    for idx, process in enumerate(sorted_processes):
        if process["processresult"] == "Successful":
            success_counter += 1
            temp = calculate_time_difference_in_minutes(process["starteddate"], process["endeddate"])
            calc_successful += temp

        elif process["processresult"] == "Faulted":
            faulted_counter += 1
            temp = calculate_time_difference_in_minutes(process["starteddate"], process["endeddate"])

        elif process["processresult"] == "Running":
            pending_counter += 1
            temp = ""
        
        if idx > 0:
            previous_process = sorted_processes[idx - 1]
            end_prev = datetime.strptime(previous_process["endeddate"], "%m/%d/%Y %I:%M:%S %p")
            start_current = datetime.strptime(process["starteddate"], "%m/%d/%Y %I:%M:%S %p")
            free_time = calculate_time_difference_in_minutes(end_prev.strftime("%m/%d/%Y %I:%M:%S %p"), start_current.strftime("%m/%d/%Y %I:%M:%S %p"))
            if end_prev < start_current and free_time >= min_free_time_minutes:
                robot_free_intervals.append({
                    "start": end_prev.strftime("%m/%d/%Y %I:%M:%S %p"),
                    "end": start_current.strftime("%m/%d/%Y %I:%M:%S %p")
                })
        
        tempcalc = calc_successful / (day_interval * 24 * 60) * 100
        tempcalc = round(tempcalc, 2)
        if type(temp) != type(""):
            temp = str(round(temp, 2)) + " dakika"
        else:
            temp = ""


        ws.append([process["name"] , process["starteddate"] , process["endeddate"] , temp , process["processresult"] , process["processstarter"]])

    utilization[robotprocesses]["efficiency"] = {
        "Total Consumed Time": str(calc_successful),
        "Successful Counter": success_counter, 
        "Faulted Counter": faulted_counter, 
        "Running Counter": pending_counter,
        "Efficiency": "%" + str(tempcalc)
    }
    utilization[robotprocesses]["available"] = robot_free_intervals

    ws.append([])
    ws.append([])


    ws.append(["Total Consumed Time", str(round(calc_successful, 3))+"dakika"])
    ws[ws.max_row][0].font = bold_font
    ws.append(["Successful Counter",success_counter])
    ws[ws.max_row][0].font = bold_font
    ws.append(["Faulted Counter",faulted_counter])
    ws[ws.max_row][0].font = bold_font
    ws.append(["Running Counter",pending_counter])
    ws[ws.max_row][0].font = bold_font
    ws.append(["Efficiency", "%" + str(tempcalc)])
    ws[ws.max_row][0].font = bold_font
    ws.append([])
    ws.append([])


    ws.append(["Available Start Times", "Available End Times","Total Free Time"])
    ws[ws.max_row][0].font = bold_font
    ws[ws.max_row][1].font = bold_font
    ws[ws.max_row][2].font = bold_font

    for gapdata in utilization[robotprocesses]["available"]:
        gap = format_time(calculate_time_difference_in_minutes(gapdata["start"],gapdata["end"]))
        ws.append([gapdata["start"], gapdata["end"], gap])


# json_output = json.dumps(utilization, ensure_ascii=False, indent=4)
# with open("utilization.json", "w", encoding="utf-8") as file:
#     file.write(json_output)


wb.remove(wb["Sheet"])
wb.save(utilizasyonraporu_cikti_dosyasi_yolu)
