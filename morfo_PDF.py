import PyPDF2
import pandas as pd
import os

filenames = next(os.walk('.'), (None, None, []))[2]

files_to_proc = []

for num in range(len(filenames)):

    if filenames[num][-4:-1]+filenames[num][-1] == '.pdf':
        files_to_proc.append(filenames[num])
    else:
        pass

###DataFrame

df = pd.DataFrame(columns = ('Examination number','Examination date','GFR','Kreatynina','BASO','BASO%','EOS','EOS%','HCT','HGB','IG',
    'IG%','LYM','LYM%','MCH','MCHC','MCV','MONO','MONO%','MPV','NEU',
    'NEU%','P-LCR','PCT','PDW','PLT','RBC','RDW-CV','RDW-SD','WBC'))

m = 0
def extraction(file_number):

    file_name = files_to_proc[file_number]

    ### Creating a pdf file object
    pdfFileObj = open(file_name, 'rb')
    # creating a pdf reader object
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    # extracting text
    total_pdf = []
    count = pdfReader.numPages
    for i in range(count):
        page = pdfReader.getPage(i)
        output = page.extractText()
        total_pdf.append(output)
    total_pdf = ' '.join(total_pdf)
    # closing the pdf file object
    pdfFileObj.close()

    ###Extracting data from string

    results = total_pdf.split('identyfikacyjny kod resortowy 000000018583-81')[1]

    patient_name = results.split(',')[0][32:]
    pesel = results.split('Badanie')[0][-11:]

    assert len(pesel) == 11

    results = total_pdf.split('Zlecenie nr ')

    for ex_number in range(1,len(results)):

        examination = results[ex_number]

        examination_number = examination.split(' z dnia ')[0]
        examination_date = examination.split(' z dnia ')[1][:10]
        
        if examination.find('GFR')>0:
            GFR = examination.split('GFR wg MDRD')[1].split('ml')[0].strip()
        else:
            GFR = 'nf'
        assert len(GFR) < 8

        if examination.find('Kreatynina')>0:    
            Kreatynina = examination.split('Kreatynina')[1].split('µmol')[0].strip()
        else:
            Kreatynina = 'nf'
        assert len(Kreatynina) < 8

        if examination.find('BASO')>0:    
            BASO = examination.split('BASO')[1].split('10´3/µl')[0].strip()
        else:
            BASO = 'nf'
        assert len(BASO) < 8

        if examination.find('BASO%')>0:    
            BASO_proc = examination.split('BASO%')[1].split('%')[0].strip()
        else:
            BASO_proc = 'nf'
        assert len(BASO_proc) < 8

        if examination.find('EOS')>0:
            EOS = examination.split('EOS')[1].split('10´3/µl')[0].strip()
        else:
            EOS = 'nf'
        assert len(EOS) < 8

        if examination.find('EOS%')>0:
            EOS_proc = examination.split('EOS%')[1].split('%')[0].strip()
        else:
            EOS_proc = 'nf'
        assert len(EOS_proc) < 8

        if examination.find('HCT  (Hematokryt)')>0:
            HCT = examination.split('HCT  (Hematokryt)')[1].split('%')[0].strip()
        else:
            HCT = 'nf'
        assert len(HCT) < 8

        if examination.find('HGB  (Hemoglobina)')>0:
            HGB = examination.split('HGB  (Hemoglobina)')[1].split('g/dl')[0].strip()
        else:
            HGB = 'nf'
        assert len(HGB) < 8

        if examination.find('IG')>0:
            IG = examination.split('IG')[1].split('10´3/µl')[0].strip()
        else:
            IG = 'nf'
        assert len(IG) < 8

        if examination.find('IG%')>0:
            IG_proc = examination.split('IG%')[1].split('%')[0].strip()
        else:
            IG_proc = 'nf'
        assert len(IG_proc) < 8

        if examination.find('LYM')>0:
            LYM = examination.split('LYM')[1].split('10´3/µl')[0].strip()
        else:
            LYM = 'nf'
        assert len(LYM) < 8

        if examination.find('LYM%')>0:
            LYM_proc = examination.split('LYM%')[1].split('%')[0].strip()
        else:
            LYM_proc = 'nf'
        assert len(LYM_proc) < 8

        if examination.find('MCH')>0:
            MCH = examination.split('MCH')[1].split('pg')[0].strip()
        else:
            MCH = 'nf'
        assert len(MCH) < 8

        if examination.find('MCHC')>0:
            MCHC = examination.split('MCHC')[1].split('g/dl')[0].strip()
        else:
            MCHC = 'nf'
        assert len(MCHC) < 8

        if examination.find('MCV')>0:
            MCV = examination.split('MCV')[1].split('fl')[0].strip()
        else:
            MCV = 'nf'
        assert len(MCV) < 8

        if examination.find('MONO')>0:
            MONO = examination.split('MONO')[1].split('10´3/µl')[0].strip()
        else:
            MONO = 'nf'
        assert len(MONO) < 8

        if examination.find('MONO%')>0:
            MONO_proc = examination.split('MONO%')[1].split('%')[0].strip()
        else:
            MONO_proc = 'nf'
        assert len(MONO_proc) < 8

        if examination.find('MPV')>0:
            MPV = examination.split('MPV')[1].split('fl')[0].strip()
        else:
            MPV = 'nf'
        assert len(MPV) < 8

        if examination.find('NEU')>0:
            NEU = examination.split('NEU')[1].split('10´3/µl')[0].strip()
        else:
            NEU = 'nf'
        assert len(NEU) < 8

        if examination.find('NEU%')>0:
            NEU_proc = examination.split('NEU%')[1].split('%')[0].strip()
        else:
            NEU_proc = 'nf'
        assert len(NEU_proc) < 8

        if examination.find('P-LCR')>0:
            P_LCR = examination.split('P-LCR')[1].split('%')[0].strip()
        else:
            P_LCR = 'nf'
        assert len(P_LCR) < 8

        if examination.find('PCT')>0:
            PCT = examination.split('PCT')[1].split('%')[0].strip()
        else:
            PCT = 'nf'
        assert len(PCT) < 8

        if examination.find('PDW')>0:
            PDW = examination.split('PDW')[1].split('fl')[0].strip()
        else:
            PDW = 'nf'
        assert len(PDW) < 8

        if examination.find('PLT  (P³ytki krwi)')>0:
            PLT = examination.split('PLT  (P³ytki krwi)')[1].split('10´3/µl')[0].strip()
        else:
            PLT = 'nf'
        assert len(PLT) < 8

        if examination.find('RBC  (Erytrocyty)')>0:	
            RBC = examination.split('RBC  (Erytrocyty)')[1].split('10´6/µl')[0].strip()	
        else:
            RBC = 'nf'
        assert len(RBC) < 8

        if examination.find('RDW-CV')>0:
            RDW_CV = examination.split('RDW-CV')[1].split('%')[0].strip()
        else:
            RDW_CV = 'nf'
        assert len(RDW_CV) < 8

        if examination.find('RDW-SD')>0:
            RDW_SD = examination.split('RDW-SD')[1].split('fl')[0].strip()
        else:
            RDW_SD = 'nf'
        assert len(RDW_SD) < 8

        if examination.find('WBC  (Leukocyty)')>0:
            WBC = examination.split('WBC  (Leukocyty)')[1].split('10´3/µl')[0].strip()
        else:
            WBC = 'nf'
        assert len(WBC) < 8

        examination_row = [examination_number,examination_date,GFR,Kreatynina,BASO,BASO_proc,EOS,EOS_proc,HCT,HGB,IG,
            IG_proc,LYM,LYM_proc,MCH,MCHC,MCV,MONO,MONO_proc,MPV,NEU,
            NEU_proc,P_LCR,PCT,PDW,PLT,RBC,RDW_CV,RDW_SD,WBC]

        df.loc[ex_number] = examination_row

    with pd.ExcelWriter(patient_name+'_'+pesel+'.xlsx',engine='xlsxwriter') as writer:  
        df.to_excel(writer)


print('######################################## PDF -> Excel ########################################')
print('Skrypt przenosi wyniki badań zapisane w formacie PDF i zapisuje je w arkuszu programu MS Excel.\n')
print('Zasady działania:')
print('1) Program składa się z jednego pliku i nie wymaga instalacji')
print('2) Analizuje wszystkie pliki w formacie PDF w folderze w którym się znajduje')
print('3) Pliki Excel zapisywane są w tym samym folderze')
print('4) Nazwa nowego pliku składa się z imienia, nazwiska i numeru PESEL z pliku PDF')
print('5) Skrypt zawiera mechanizm samosprawdzający: pojawienie się komunikatu AssertionError\noznacza że plik PDF nie pasuje do szablonu')
print('\nby Ipliph\n')
input('Wciśnij Enter aby kontynuować...')

m = '1'
for num in range(len(files_to_proc)):
    print(files_to_proc[num]+' jest analizowany...('+m+'/'+str(len(files_to_proc))+')')
    extraction(num)
    m = str(int(m) + 1)  

print('\nWszystkie znalezione pliki zostały przeliczone')
input('Wciśnij Enter aby zakończyć...')