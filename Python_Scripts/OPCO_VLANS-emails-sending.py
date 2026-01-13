import os
import win32com.client as win32
import time

FOLDER_PATH = r"C:\Users\lynette.mutuku\Downloads\November"

recipient_map = {
    'MTN Benin': '"Aristide Maffon [ MTN Benin ]" <Aristide.Maffon@mtn.com>; "Cherif Bachabi [ MTN Benin ]" <Cherif.Bachabi@mtn.com>; '
    '"Jonas Martin [ MTN Benin ]" <Jonas.Martin@mtn.com>; "Melissa Tokouete [ MTN Benin ]" <Melissa.Tokouete@mtn.com>; '
    '"Romic DENON [ MTN Benin ]" <Romic.Denon@mtn.com>; "sadiq.yahaya.bj@ieng-group.com" <sadiq.yahaya.bj@ieng-group.com>; '
    '"lgandonou.bj@ieng-group.com" <lgandonou.bj@ieng-group.com>; "drosius.bj@ieng-group.com" <drosius.bj@ieng-group.com>;'
    ' "sdumassi.bj@ieng-group.com" <sdumassi.bj@ieng-group.com>; "achaka.bj@ieng-group.com" <achaka.bj@ieng-group.com>;'
    ' "jandoh.bj@ieng-group.com" <jandoh.bj@ieng-group.com>; "lmaduka.bj@ieng-group.com" <lmaduka.bj@ieng-group.com>; '
    '"enwosu.bj@ieng-group.com" <enwosu.bj@ieng-group.com>; "MUKTAR MUSA [ MTN Benin ]" <Muktar.Musa@mtn.com>; '
    '"RAMSES SOGLO [ MTN Benin ]" <Ramses.Soglo@mtn.com>; "Frejus FADONOUGBO [ MTN Benin ]" <Frejus.Fadonougbo@mtn.com>; '
    '"Marius SAH [ MTN Benin ]" <Marius.Sah@mtn.com>; "hgupta.bj@ieng-group.com" <hgupta.bj@ieng-group.com>; '
    '"pakul.bj@ieng-group.com" <pakul.bj@ieng-group.com>; "agbafa.bj@ieng-group.com" <agbafa.bj@ieng-group.com>; '
    '"fkinful.bj@ieng-group.com" <fkinful.bj@ieng-group.com>; "nlaryea.bj@ieng-group.com" <nlaryea.bj@ieng-group.com>; '
    '"radetanmi.bj@ieng-group.com" <radetanmi.bj@ieng-group.com>; "ztaj@ieng-group.com" <ztaj@ieng-group.com>; '
    '"muhammad.rashid@ieng-group.com" <muhammad.rashid@ieng-group.com>; "Minhaj Ud Din Chughtai" <mchughtai@ieng-group.com>; '
    '"uashraf@ieng-group.com" <uashraf@ieng-group.com>; "alatif@ieng-group.com" <alatif@ieng-group.com>; '
    '"Ehiabhi ENIMAKPOKPO [ MTN Benin ]" <Ehiabhi.Enimakpokpo@mtn.com>; "Jerome ASSIMA [ MTN Benin ]" <jerome.assima@mtn.com>;'
    ' "Gildas Hounsou [ MTN Benin ]" <Gildas.Hounsou@mtn.com>',

    'MTN Botswana': '"Titose Isaac  [ MTN Botswana ]" <Titose.Isaac@mtn.com>; "One Motladiile [ MTN Botswana ]" <One.Motladiile@mtn.com>',

    'MTN Cameroon': "Romeo.Jipap@mtn.com; Bertin.Ngatia@mtn.com; Ebot.Daniel@mtn.com; Patrice.Ndo@mtn.com; Josiane.Ngoh@mtn.com; sandrine.gwanwo@mtn.com; Albert.Essomba@mtn.com;  Emmanuel.Ndombo@mtn.com; Aline.Jounewe@mtn.com; Walters.Tangwe@mtn.com; olivier.sipanko@mtn.com; arnaud.t@huawei.com; Aboubakar.Kouotou@mtn.com; Franck.Kom@mtn.com; faustin.tchassem@mtn.com; moumie.yves@huawei-partners.com; sefindjim.alexis.valere2@huawei-partners.com; richard.kokossa@mtn.com; Tanuequddus.Atanga@mtn.com; nyah.wx1229907@huawei-partners.com",
    
    'MTN CIV': "Maxime.Tra@mtn.com; Charles.Nossou@mtn.com; Celestin.Saraka@mtn.com; Bi.IRIE@mtn.com; Daniel.Koffi@mtn.com; Seydou.Diabate@mtn.com; Samuel.Lorougnon@mtn.com; Jonathan.NIAZALE@mtn.com; Serge.Konan@mtn.com; Andre.Beugre@mtn.com; Alain.Diby@mtn.com; Michael.FobiFokou@mtn.com; Stephane.Del_Grao@mtn.com; Daniel.Fossou@mtn.com; traoreb.adama@huawei-partners.com; yapo.adon.hubert@huawei-partners.com; diomande.blonde1@huawei-partners.com; yachy.ferdinand@huawei-partners.com; Hermann.Kouassi@mtn.com; Jean-Claude.Siaba@mtn.com;'ngbesso.aka.xavier@huawei-partners.com; ngbesso.aka.xavier@huawei-partners.com",

    'MTN Congo B': '"Farai CHUCHU [ MTN Congo Brazzaville ]" <Farai.Chuchu@mtn.com>; katuri.tawanda <katuri.tawanda@zte.com.cn>; '
    'prince.bhurabhura@zte.com.cn <prince.bhurabhura@zte.com.cn>; nguimby.clevy <nguimby.clevy@zte.com.cn>; '
    'morris.nimubona <morris.nimubona@zte.com.cn>; loulendo.wilfried1@zte.com.cn <loulendo.wilfried1@zte.com.cn>;'
    ' "Nafed FINA [ MTN Congo Brazzaville ]" <Nafed.FINA@mtn.com>; taurai.chimusimbe <taurai.chimusimbe@zte.com.cn>;'
    ' "Ondze BOCKALYCO [ MTN Congo Brazzaville ]" <Ondze.BOCKALYCO@mtn.com>; stephen.omoto <stephen.omoto@zte.com.cn>;'
    ' "Alain NGENDAHAYO [ MTN Congo Brazzaville ]" <Alain.NGENDAHAYO@mtn.com>; "Farial DIKOBAT [ MTN Congo Brazzaville ]" <Farial.DIKOBAT@mtn.com>; '
    'bassonguila.serge@zte.com.cn <bassonguila.serge@zte.com.cn>',

    'MTN Eswatini': '"Mncedisi Bhembe [ MTN Eswatini ]" <Mncedisi.Bhembe@mtn.com>; '
    '"Sihle Mahlalela [ MTN Eswatini ]" <Sihle.Mahlalela@mtn.com>; "Mlandvo Simelane [ MTN Eswatini ]" <Mlandvo.Simelane@mtn.com>;'
    ' "Makhosini Shongwe [ MTN Eswatini ]" <Makhosini.Shongwe@mtn.com>; "Nomathemba Cele [ MTN Eswatini ]" <Nomathemba.Cele@mtn.com>; '
    'nkosivile.msibi <nkosivile.msibi@huawei.com>; "Bongukuphila Mhlanga [ MTN Eswatini ]" <Bongukuphila.Mhlanga@mtn.com>; '
    '"Senzo Vilakati [ MTN Eswatini ]" <Senzo.Vilakati@mtn.com>; "Nosipho Mamba [ MTN Eswatini ]" <Nosipho.Mamba@mtn.com>',

    'MTN Ghana': '"William Awuku-Asabere [ MTN Ghana ]" <William.Awuku-Asabere@mtn.com>; '
    '"Christian Nyame [ MTN Ghana ]" <Christian.Nyame@mtn.com>; "Ibrahim Musah [ MTN Ghana ]" <Ibrahim.Musah@mtn.com>;'
    ' "Asriel Narh [ MTN Ghana ]" <Asriel.Narh@mtn.com>; "Felix Osakonor [ MTN Ghana ]" <Felix.Osakonor@mtn.com>; '
    '"Bernard Lemawu [ MTN Ghana ]" <Bernard.Lemawu@mtn.com>; "Amda Kwapong [ MTN Ghana ]" <Amda.Kwapong@mtn.com>;'
    ' "Abdul C Nasser [ MTN Ghana ]" <Abdul.CNasser@mtn.com>; "Michael Ghartey [ MTN Ghana ]" <Michael.Ghartey@mtn.com>; '
    '"Michael Kwankye [ MTN Ghana ]" <Michael.Kwankye@mtn.com>; "Jerry Nyamekye [ MTN Ghana ]" <Jerry.Nyamekye@mtn.com>; '
    '"Michael Andoh [ MTN Ghana ]" <Michael.Andoh@mtn.com>; "Alfred Addo [ MTN Ghana ]" <Alfred.Addo@mtn.com>; '
    '"Benjamin Zato [ MTN Ghana ]" <Benjamin.Zato@mtn.com>; "Pascal Dawson [ MTN Ghana ]" <Pascal.Dawson@mtn.com>;'
    ' "Bishop Ahene-Truly [ MTN Ghana ]" <Bishop.Ahene-Truly@mtn.com>; "Edward Anyidoho [ MTN Ghana ]" <Edward.Anyidoho@mtn.com>',

    # 'MTN Guinea Conakry': '"Mamadouba Sylla [ MTN Guinea-Republic ]" <Mamadouba.Sylla@mtn.com>; '
    # '"Amadou Baldé [ MTN Guinea-Republic ]" <AmadouOury.Balde@mtn.com>; "Abdourahamane Barry [ MTN Guinea-Republic ]" <Abdourahamane.Barry@mtn.com>; '
    # '"Emile Feindouno [ MTN Guinea-Republic ]" <Emile.Feindouno@mtn.com>; "Julio De Souza [ MTN Guinea-Republic ]" <Julio.DeSouza@mtn.com>; '
    # '"Aboubacar Barry [ MTN Guinea-Republic ]" <Aboubacar.Barry1@mtn.com>; '
    # '"Mawa Doumbouya [ MTN Guinea-Republic ]" <Mawa.Doumbouya@mtn.com>; "Mamandi Sidime [ MTN Guinea-Republic ]" <Mamandi.Sidime@mtn.com>',

    'MTN Liberia': "Damein.Mariappen@mtn.com; Yewande.Olaiya-Oni@mtn.com; Sonya.Monger@mtn.com; Gaston.Endeme@mtn.com; Kolubah.Flomo2@mtn.com; Mohammed.Adams@mtn.com; Jaygbah.Massaquoi@mtn.com; Ritan.Bhandari@mtn.com; Dovert.Andrews@mtn.com; anthony.n.barshea@huawei.com; Clementine.Norris@mtn.com; Joseph.Tellewoyan@mtn.com; Jitendra.Singh@mtn.com; Karl.Logan@mtn.com; Rohit.Narayan@mtn.com; Gabriel.Bedell@mtn.com; saeed.mohammed@huawei.com; siaka.bai.kpaka@huawei.com; reuben.k.gbelly@huawei.com; melvin.ed.jr@huawei.com; dovert.d.andrews@huawei.com; alphonso.manneh.sr@huawei.com; david.p.kpadeh1@h-partners.com",
    
    'MTN Nigeria': '"Olukemi Molade [MTN Nigeria]" <olukemi.molade1@mtn.com>; "Adewale Sulaiman [ MTN Nigeria ]" <Adewale.Sulaiman@mtn.com>; '
    '"David Onoja [ MTN Nigeria ]" <David.Onoja@mtn.com>; "Olatunde Olatunbosun [ MTN Nigeria ]" <Olatunde.Olatunbosun@mtn.com>; '
    '"Kelly Osagie [ MTN Nigeria ]" <Kelly.Osagie@mtn.com>; "Aaron Ogheze [ MTN Nigeria - EB ]" <Aaron.Ogheze@mtn.com>;'
    ' "Adebowale Adeshina [ MTN Nigeria ]" <Adebowale.Adeshina@mtn.com>; "Benjamin Adereti [ MTN Nigeria ]" <Benjamin.Adereti@mtn.com>;'
    ' "Tayo Onaolapo [ MTN Nigeria ]" <Tayo.Onaolapo@mtn.com>; "Olanrewaju Akinlade [ MTN Nigeria ]" <Olanrewaju.Akinlade@mtn.com>; '
    '"Elisha Adamu [ MTN Nigeria ]" <Elisha.Adamu@mtn.com>; "Adefolarin Ajakaiye_Old [ MTN Nigeria ]" <Adefolarin.Ajakaiye@mtn.com>;'
    ' "Adeyemi Adeola [ MTN Nigeria ]" <Adeyemi.Adeola@mtn.com>; "Adebisi Adewole [ MTN Nigeria ]" <Adebisi.Adewole@mtn.com>; '
    '"Oluwafemi Yakubu [ MTN Nigeria ]" <Oluwafemi.Yakubu@mtn.com>; "Adeola Alade [ MTN Nigeria ]" <Adeola.Alade@mtn.com>; '
    '"Akinwale Fabumuyi [MTN Nigeria]" <Akinwale.Fabumuyi1@mtn.com>; "Emeka Nwaosa [ MTN Nigeria ]" <emeka.nwaosa@mtn.com>; '
    '"Etiene Moses [ MTN Nigeria ]" <Etiene.Moses@mtn.com>; "Oluwaremilekun Odiyi [ MTN Nigeria ]" <Oluwaremilekun.Odiyi@mtn.com>; '
    '"Korede Ayoola [ MTN Nigeria ]" <Korede.Ayoola@mtn.com>; "Adebayo Obayomi [ MTN Nigeria ]" <Adebayo.Obayomi@mtn.com>',


    'MTN Rwanda': '"Issa Nkusi Karera [ MTN Rwanda ]" <Issa.NkusiKarera@mtn.com>; "Nicolas Kimana [ MTN Rwanda ]" <Nicolas.Kimana@mtn.com>;'
    ' "Justin Rugamba [ MTN Rwanda ]" <Justin.Rugamba@mtn.com>; "Rene Manzi [ MTN Rwanda ]" <Rene.Manzi@mtn.com>; '
    '"Betty Dusabe [ MTN Rwanda ]" <Betty.Dusabe@mtn.com>; "Job Munyaneza [ MTN Rwanda ]" <Job.Munyaneza@mtn.com>; '
    '"Claire Kayitesi [ MTN Rwanda ]" <Claire.Kayitesi@mtn.com>; "Viateur Mugenzi [ MTN Rwanda ]" <Viateur.Mugenzi@mtn.com>; '
    '"Vivy Ingabire [ MTN Rwanda ]" <Vivy.Ingabire@mtn.com>; "Pacifique Tuyishime [ MTN Rwanda ]" <Pacifique.Tuyishime@mtn.com>;'
    ' "Sarah Umutoni [ MTN Rwanda ]" <Sarah.Umutoni@mtn.com>; "Rodney NIYIKORA [ MTN Rwanda ]" <Rodney.Niyikora@mtn.com>;'
    ' "Aimable Mwizerwa [ MTN Rwanda ]" <Aimable.MwizerwaDady@mtn.com>; "Bonck Ngabo [ MTN Rwanda ]" <Bonck.Ngabo@mtn.com>; '
    '"Clement NIYIGIRIMBABAZI [ MTN Rwanda ]" <Clement.Niyigirimbabazi@mtn.com>; "Back Office" <back_office.rw@mtn.com>; '
    'nshuti patrick <nshuti.patrick@ericsson.com>; emmanuel mbonigaba <emmanuel.mbonigaba@ericsson.com>;'
    ' theogene niyonzima <theogene.niyonzima@ericsson.com>',

    'MTN South Africa': '"Thulani Mhlanga [ MTN South Africa ]" <Thulani.Mhlanga@mtn.com>; "Zaheer Coovadia [ MTN South Africa ]" <Zaheer.Coovadia@mtn.com>; '
    '"Gideon Maggerman [ MTN South Africa ]" <Gideon.Maggerman@mtn.com>; '
    '"Patricia Mokoena [ MTN SA - African Ideas Corporation (Pty) Ltd ]" <Patricia.Mokoena@mtn.com>; "Ntswaki Possa [ MTN South Africa ]" <Ntswaki.Possa@mtn.com>;'
    ' "Ashley Jacobs [ MTN South Africa ]" <Ashley.Jacobs@mtn.com>; "Kitchner Malepe [ MTN SA - African Ideas Corporation (Pty) Ltd ]" <Kitchner.Malepe@mtn.com>;'
    ' "Sibusisiwe Bujela [ MTN SA - AFRICAN IDEAS CORPORATION PTY LTD ]" <Sibusisiwe.Bujela@mtn.com>; '
    '"Abri Potgieter [ MTN South Africa ]" <Abri.Potgieter@mtn.com>; "Dharmesh Kassen [ MTN South Africa ]" <Dharmesh.Kassen@mtn.com>;'
    ' "Bulelani Kalolo [ MTN South Africa ]" <Bulelani.Kalolo@mtn.com>; "Thandiswa Siguca [ MTN South Africa ]" <Thandiswa.Siguca@mtn.com>; '
    '"Paulina Tsengiwe [ MTN South Africa ]" <Paulina.Tsengiwe@mtn.com>',

    'MTN South Sudan': '"Brian Karuhanga [ MTN South Sudan ]" <Brian.Karuhanga@mtn.com>; "Nathaniel Seku [ MTN South Sudan ]" <Nathaniel.Seku@mtn.com>;'
    ' "Joshua Nomwesigwa [ MTN South Sudan ]" <Joshua.Nomwesigwa@mtn.com>; "James Mugumire" <jmugumire@ieng-group.com>; "Yamin Hassan [ MTN South Sudan ]" <Yamin.Hassan@mtn.com>; '
    '"Kipruto Keter [ MTN South Sudan ]" <Kipruto.Keter@mtn.com>; "Victor Ssegawa [ MTN South Sudan ]" <Victor.Ssegawa@mtn.com>; '
    '"Muhammad Asim [ MTN South Sudan ]" <Muhammad.Asim@mtn.com>; "Augustino Dany [ MTN South Sudan ]" <Augustino.Dany@mtn.com>; '
    '"James Ogola [ MTN South Sudan ]" <James.Ogola@mtn.com>; "Muhammad Adeel Riaz" <mriaz@ieng-group.com>; "Moris Henry [ MTN South Sudan ]" <Moris.Henry@mtn.com>; '
    '"John Banga" <Jbanga@ieng-group.com>; "Samar Musa Almak" <salmak@ieng-group.com>; "musab.ahmed@ieng-group.com" <musab.ahmed@ieng-group.com>;'
    ' "smutai@ieng-group.com" <smutai@ieng-group.com>; "Khidr Ahmed" <kahmed@ieng-group.com>; "Ayesha Farrukh" <afarrukh@ieng-group.com>; "Anthoney  Dabi" <adabi@ieng-group.com>',

    'MTN Uganda': '"Philip Wamimbi [ MTN Uganda ]" <Philip.Wamimbi@mtn.com>; "Carol Akimanzi [ MTN Uganda ]" <Carol.Akimanzi@mtn.com>; '
    '"Joseph Sakwa [ MTN Uganda ]" <Joseph.Sakwa@mtn.com>; "Stephen Mubiru [ MTN Uganda ]" <Stephen.Mubiru@mtn.com>;'
    ' "Robbins Mwehair [ MTN Uganda ]" <Robbins.Mwehair@mtn.com>; "Samuel Sentongo [ MTN Uganda ]" <Sentongo.Sentongo@mtn.com>;'
    ' "Mustapha Kagumba [ MTN Uganda ]" <Mustapha.Kagumba@mtn.com>; "Collin Kiwotoka [ MTN Uganda ]" <Collin.Kiwotoka@mtn.com>; '
    '"Peninah Nankya [ MTN Uganda ]" <Peninah.Nankya@mtn.com>; "Desta Chemery [ MTN Uganda ]" <Desta.Chemery@mtn.com>;'
    ' "Fred Sekyana [ MTN Uganda ]" <Fred.Sekyana@mtn.com>; "Joshua Sabiti [ MTN Uganda ]" <Joshua.Sabiti@mtn.com>;'
    ' "Stephen Mugisha [ MTN Uganda ]" <Stephen.Mugisha@mtn.com>; "Edison Ganati [ MTN Uganda ]" <Edison.Ganati@mtn.com>; '
    '"Enock Mwerinde [ MTN Uganda ]" <Enock.Mwerinde@mtn.com>; "Geraldine Mwika [ MTN Uganda ]" <Geraldine.Mwika@mtn.com>',
   
    'MTN Zambia': "Wilson.Nkunika@mtn.com; Martin.Mutanya@mtn.com>; joseph.mufaya@huawei.com; Precious.Mutale@mtn.com;  Kanchule.Sinyangwe@mtn.com; winston.mtonga1@huawei.com; Patience.Banda@mtn.com; Yowela.Katakwe@mtn.com; Mumbelunga.Funjika@mtn.com; albert.masialeti@huawei.com; enock.muyambo1@huawei-partners.com; likando.singongi@huawei.com; Isaac.Chibanga@Bayobab.Africa; Chafika.Mwale@mtn.com; David.Daka@mtn.com; Nanji.Namusamba@mtn.com; Lopi.mutati@huawei.com; Kaluwe.Moonde@mtn.com; bwalya.augustus@huawei.com; gideon.mfula@huawei.com; Makiyo.Muwirimi1@mtn.com; Jacob.Taingishi@mtn.com",

    
}

cc_email = "Anthony.Chintoh@Bayobab.Africa; Luc-Fabrice.Ndifor@Bayobab.Africa; Gilbert.Sang@Bayobab.Africa; Michael.Sabala@Bayobab.Africa; Craig.Beacham@Bayobab.Africa; martin.nganga@bayobab.africa; SR-F@bayobab.africa; dennis.cheruiyot@bayobab.africa; Lynette.Mutuku@Bayobab.Africa;"


# ✉️ Email subject and body
email_subject = 'Utilization Report_November 2025'
email_body = '''
<div style="font-family: Poppins, sans-serif; font-size:10pt; color:#44536a;">

  <p>Dear Team,</p>

  <p>Please find attached the Utilization Report for November 2025.</p>

  <p>Regards,</p>
</div>
'''

outlook = win32.Dispatch('Outlook.Application')

# accounts = outlook.Session.Accounts
# for account in accounts:
#     if account.SmtpAddress.lower() == "test@mutuku.com":
#         mail.SendUsingAccount = account
#         print(f"📤 Sending from: {account.SmtpAddress}")
#         break
# else:
#     print("⚠️ Could not find sender account: test@mutuku.com")

for file in os.listdir(FOLDER_PATH):
    if file.endswith('.pdf'):
        for identifier, to_email in recipient_map.items():
            print("identifier", identifier)
            if file.startswith(identifier):
                print(f"✅ Matched file '{file}' to recipient '{to_email}'")

                full_path = os.path.join(FOLDER_PATH, file)

                # Create new email
                mail = outlook.CreateItem(0)
                mail.To = to_email
                mail.CC = cc_email
                mail.Subject = f'{identifier} - {email_subject}'

                for account in outlook.Session.Accounts:
                    if account.SmtpAddress.lower() == "QE-F@bayobab.Africa":
                        mail.SendUsingAccount = account
                        print(f"📤 Sending from: {account.SmtpAddress}")
                        break
                else:
                    print("⚠️ Could not find sender account: QE-F@bayobab.africa")

                # Open to insert signature
                mail.Display()
                time.sleep(1)

                # Prepend custom body to the existing signature
                original_body = mail.HTMLBody
                mail.HTMLBody = email_body + original_body

                # Attach the file
                mail.Attachments.Add(full_path)

                print(f'📬 Opened draft for {to_email} with attachment: {file}')
                break
        else:
            print(f'⚠️ No matching recipient found for file: {file}')


# for file in os.listdir(FOLDER_PATH):
#     if file.endswith('.pdf'):
#         for identifier, to_email in recipient_map.items():
#             if file.startswith(identifier):
#                 print(f"✅ Matched file '{file}' to recipient '{to_email}'")

#                 full_path = os.path.join(FOLDER_PATH, file)

#                 # Create new email
#                 mail = outlook.CreateItem(0)
#                 mail.To = to_email
#                 # mail.Subject = email_subject
#                 mail.Subject = f'{identifier} - {email_subject}'


#                 # Open to insert signature
#                 mail.Display()
#                 time.sleep(1)

#                 # Prepend custom body to the existing signature
#                 original_body = mail.HTMLBody
#                 mail.HTMLBody = email_body + original_body

#                 # Attach the file
#                 mail.Attachments.Add(full_path)

#                 print(f'📬 Opened draft for {to_email} with attachment: {file}')
#                 break  # Move to next file after finding a match
#         else:
#             print(f'⚠️ No matching recipient found for file: {file}')