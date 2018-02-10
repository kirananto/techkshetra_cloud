const functions = require('firebase-functions')
const admin = require('firebase-admin')
const nodemailer = require('nodemailer')
const Datauri = require('datauri')
const datauri = new Datauri()
const CryptoJS = require("crypto-js")
const QRCode =  require('qrcode')
const Excel = require('exceljs')
admin.initializeApp(functions.config().firebase)
var counter = 1
const firestore = admin.firestore()
const gmailEmail = 'kirananto@gmail.com';
var bytes  = CryptoJS.AES.decrypt(ciphertext, 'password');
var gmailPassword = bytes.toString(CryptoJS.enc.Utf8);
const mailTransport = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: gmailEmail,
    pass: gmailPassword
  }
})
const mailOptions = {
  from: '"Techkshetra" <kirananto@gmail.com>',
  subject: 'Thanks for registering for Techkshetra18',
  to:  'kirananto@gmail.com,kirananto@hotmail.com,lijopaul1996@gmail.com,ajmalv31@gmail.com,josephvarghese.rms@gmail.com'
}

exports.registered = functions.firestore
  .document('registration/{userId}/registered/{eventId}')
  .onWrite(event => {
    var newValue = event.data.data()
    newValue.eventId = event.params.eventId
    QRCode.toDataURL(event.params.userId, { errorCorrectionLevel: 'H' })
            .then(url => {
              
                mailOptions.to = newValue.participants[0].email
                mailOptions.html = `<h1>Thanks for registering.</h1>
                <h2> You have Registered for ${newValue.eventId} </h2>
                <h2> Please bring the qrcode for payment</h2>
                <h2> Also don't forget to bring your college issued id card without
                 which you will not be admitted for the event </h2>
                 <img src="cid:unique@kreata.ee"/>`
                mailOptions.attachments = [{
                  path: url,
                  cid: 'unique@kreata.ee'
                }]
                return mailTransport.sendMail(mailOptions)
                  .then(() => {
                    console.log(`New confirmation email sent to:`, newValue.participants[0].email)
                    var registeredRef = firestore.collection('registered/').add(newValue)
                    .then(success => {
                      console.log('success')
                    })
                  })
                  .catch(error => console.error('There was an error while sending the email:', error));
            })
            .catch(err => {
                console.error(err)
                return
            })
})  

exports.workshopsPaidList = functions.firestore.document('export/workshoppaid').onWrite(event => {
  mailOptions.subject = 'Techkshetra 18 - Workshops paid list '
  mailOptions.html = '<h1> Complete Workshops Paid List </h1>'
  mailOptions.attachments = []
  return registeredRef = firestore.collection('workshops').get().then(querySnapshot => {
    var workbook = new Excel.Workbook()
    workbook.creator = 'Kiran Anto'
    workbook.created = new Date()
    workbook.modified = new Date()
    workbook.views = [
      {
        x: 0, y: 0, width: 100, height: 100,
        firstSheet: 0, activeTab: 1, visibility: 'visible'
      }
    ]
    var p1 = new Promise((resolve,reject) => {
      counter = querySnapshot.size
      querySnapshot.forEach(doc => { 
        firestore.collection('paid').where('eventId', '==', doc.id).get().then(query => {
          counter = counter -1
          if (query.size > 0) {
            var sheet = workbook.addWorksheet(doc.id)
            sheet.addRow(['UID', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE'])
            sheet.getRow(1).font = { bold: true}
            var total_reg = 0
            query.forEach(docum => {
                var rows = []
                rows.push(docum.data().uid)
                docum.data().participants.forEach(participant => {
                  rows.push(participant.displayName, participant.email, participant.mobno, participant.college)
                  total_reg++
                })
                sheet.addRow(rows)
            })
            sheet.addRow([])
            sheet.addRow(['','','TOTAL REGISTRATIONS :', `${total_reg}`])
          }
          if ( counter === 0) {
            resolve()
          }
        })
      })
    }).then(() => {
      workbook.xlsx.writeBuffer({
            base64: true
          }).then(buffer => {
            datauri.format('.xlsx', buffer)
            mailOptions.attachments.push({
              path: datauri.content,
              filename: 'Workshops Paid List'
            })
            mailTransport.sendMail(mailOptions)
              .then(() => console.log(`Mail send with ${mailOptions.attachments.length} attachments`))
              .catch(error => console.error('There was an error while sending the email:', error))
          })
    })
  })
})

exports.totalCollection = functions.firestore.document('export/total').onWrite(event => {
  mailOptions.subject = 'Techkshetra 18 - Total Collections '
  return colRef = firestore.collection('cashiers').doc('todays').getCollections().then(collection => {
    let counter = 0
    console.log(collection.length)
    let p1 = new Promise((resolve, reject) => {
      collection.forEach(querySnapshot => {
        firestore.collection('cashiers').doc('todays').collection(querySnapshot.id).get().then(query => {
          var total = 0
          query.forEach(doc => {
            total+=parseInt(doc.data().amount)
          })
          counter++
          mailOptions.html = `<h2> ${querySnapshot.id}  -  ${total} Rupees </h2>`
          console.log(counter)
          if (counter === collection.length) {
            resolve()
          }
        })
      })
    })
    p1.then(success => {
      mailTransport.sendMail(mailOptions)
              .then(() => console.log(`Mail send with amount details`))
              .catch(error => console.error('There was an error while sending the email:', error))
    })
  })
})

exports.eventPaidList = functions.firestore.document('export/eventpaid').onWrite(event => {
  mailOptions.subject = 'Techkshetra 18 - Events paid list '
  mailOptions.html = '<h1> Complete Event Paid List </h1>'
  mailOptions.attachments = []
  return registeredRef = firestore.collection('events').get().then(querySnapshot => {
    var workbook = new Excel.Workbook()
    workbook.creator = 'Kiran Anto'
    workbook.created = new Date()
    workbook.modified = new Date()
    workbook.views = [
      {
        x: 0, y: 0, width: 100, height: 100,
        firstSheet: 0, activeTab: 1, visibility: 'visible'
      }
    ]
    var p1 = new Promise((resolve,reject) => {
      counter = querySnapshot.size
      querySnapshot.forEach(doc => { 
        firestore.collection('paid').where('eventId', '==', doc.id).get().then(query => {
          counter = counter -1
          if (query.size > 0) {
            var sheet = workbook.addWorksheet(doc.id)
            sheet.addRow(['UID', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE'])
            sheet.getRow(1).font = { bold: true}
            var total_reg = 0
            query.forEach(docum => {
                var rows = []
                rows.push(docum.data().uid)
                docum.data().participants.forEach(participant => {
                  rows.push(participant.displayName, participant.email, participant.mobno, participant.college)
                  total_reg++
                })
                sheet.addRow(rows)
            })
            sheet.addRow([])
            sheet.addRow(['','','TOTAL REGISTRATIONS :', `${total_reg}`])
          }
          if ( counter === 0) {
            resolve()
          }
        })
      })
    }).then(() => {
      workbook.xlsx.writeBuffer({
            base64: true
          }).then(buffer => {
            datauri.format('.xlsx', buffer)
            mailOptions.attachments.push({
              path: datauri.content,
              filename: 'Events Paid List'
            })
            mailTransport.sendMail(mailOptions)
              .then(() => console.log(`Mail send with ${mailOptions.attachments.length} attachments`))
              .catch(error => console.error('There was an error while sending the email:', error))
          })
    })
  })
})


exports.workshopsList = functions.firestore.document('export/workshop').onWrite(event => {
  mailOptions.subject = 'Techkshetra 18 - Workshops list '
  mailOptions.html = '<h1> Complete Workshops List </h1>'
  mailOptions.attachments = []
  return registeredRef = firestore.collection('workshops').get().then(querySnapshot => {
    var workbook = new Excel.Workbook()
    workbook.creator = 'Kiran Anto'
    workbook.created = new Date()
    workbook.modified = new Date()
    workbook.views = [
      {
        x: 0, y: 0, width: 100, height: 100,
        firstSheet: 0, activeTab: 1, visibility: 'visible'
      }
    ]
    var p1 = new Promise((resolve,reject) => {
      counter = querySnapshot.size
      querySnapshot.forEach(doc => { 
        firestore.collection('registered').where('eventId', '==', doc.id).get().then(query => {
          counter = counter -1
          if (query.size > 0) {
            var sheet = workbook.addWorksheet(doc.id)
            sheet.addRow(['UID', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE'])
            sheet.getRow(1).font = { bold: true}
            var total_reg = 0
            query.forEach(docum => {
                var rows = []
                rows.push(docum.data().uid)
                docum.data().participants.forEach(participant => {
                  rows.push(participant.displayName, participant.email, participant.mobno, participant.college)
                  total_reg++
                })
                sheet.addRow(rows)
            })
            sheet.addRow([])
            sheet.addRow(['','','TOTAL REGISTRATIONS :', `${total_reg}`])
          }
          if ( counter === 0) {
            resolve()
          }
        })
      })
    }).then(() => {
      workbook.xlsx.writeBuffer({
            base64: true
          }).then(buffer => {
            datauri.format('.xlsx', buffer)
            mailOptions.attachments.push({
              path: datauri.content,
              filename: 'Workshops List'
            })
            mailTransport.sendMail(mailOptions)
              .then(() => console.log(`Mail send with ${mailOptions.attachments.length} attachments`))
              .catch(error => console.error('There was an error while sending the email:', error))
          })
    })
  })
})

exports.eventList = functions.firestore.document('export/event').onWrite(event => {
  mailOptions.subject = 'Techkshetra 18 - Events list '
  mailOptions.html = '<h1> Complete Event List </h1>'
  mailOptions.attachments = []
  return registeredRef = firestore.collection('events').get().then(querySnapshot => {
    var workbook = new Excel.Workbook()
    workbook.creator = 'Kiran Anto'
    workbook.created = new Date()
    workbook.modified = new Date()
    workbook.views = [
      {
        x: 0, y: 0, width: 100, height: 100,
        firstSheet: 0, activeTab: 1, visibility: 'visible'
      }
    ]
    var p1 = new Promise((resolve,reject) => {
      counter = querySnapshot.size
      querySnapshot.forEach(doc => { 
        firestore.collection('registered').where('eventId', '==', doc.id).get().then(query => {
          counter = counter -1
          if (query.size > 0) {
            var sheet = workbook.addWorksheet(doc.id)
            sheet.addRow(['UID', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE', 'PARTICIPANT', 'EMAIL', 'PH.NO', 'COLLEGE'])
            sheet.getRow(1).font = { bold: true}
            var total_reg = 0
            query.forEach(docum => {
                var rows = []
                rows.push(docum.data().uid)
                docum.data().participants.forEach(participant => {
                  rows.push(participant.displayName, participant.email, participant.mobno, participant.college)
                  total_reg++
                })
                sheet.addRow(rows)
            })
            sheet.addRow([])
            sheet.addRow(['','','TOTAL REGISTRATIONS :', `${total_reg}`])
          }
          if ( counter === 0) {
            resolve()
          }
        })
      })
    }).then(() => {
      workbook.xlsx.writeBuffer({
            base64: true
          }).then(buffer => {
            datauri.format('.xlsx', buffer)
            mailOptions.attachments.push({
              path: datauri.content,
              filename: 'Events List'
            })
            mailTransport.sendMail(mailOptions)
              .then(() => console.log(`Mail send with ${mailOptions.attachments.length} attachments`))
              .catch(error => console.error('There was an error while sending the email:', error))
          })
    })
  })
})
