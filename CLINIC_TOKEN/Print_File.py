
def create_text_file(name , token ,date , dr_name, patient_id , consultation_charge):
    f = open('Token_to_Print.txt', 'a+')
    f.truncate(0)
    f.write('*'*30 +" Consultation Token "+'*'*30 +'\n'+
            'Date : '+ date + ' '*40 + 'Token Number : '+ token +'\n\n'+
            'Patient ID   : '+patient_id+'\n\n'+
            'Patient Name : ' + name + '\n\n' +
            "Doctor's Name : "+ dr_name +'\n\n' +
            "Connsultation Fees : Rs."+consultation_charge +'\n\n' +
            "*"*80
            )
    f.close()