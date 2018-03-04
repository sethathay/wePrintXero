'use strict';

const express = require('express');
const session = require('express-session');
const xero = require('xero-node');
const exphbs = require('express-handlebars');
const fs = require('fs');
var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');
var path = require('path');
var dateFormat = require('dateformat');
var numeral = require('numeral');
var msopdf = require('node-msoffice-pdf');

var xeroClient;
var eventReceiver;
var metaConfig = {};

var app = express();

var exbhbsEngine = exphbs.create({
    defaultLayout: 'main',
    layoutsDir: __dirname + '/views/layouts',
    partialsDir: [
        __dirname + '/views/partials/'
    ],
    helpers: {
        ifCond: function(v1, operator, v2, options) {

            switch (operator) {
                case '==':
                    return (v1 == v2) ? options.fn(this) : options.inverse(this);
                case '===':
                    return (v1 === v2) ? options.fn(this) : options.inverse(this);
                case '!=':
                    return (v1 != v2) ? options.fn(this) : options.inverse(this);
                case '!==':
                    return (v1 !== v2) ? options.fn(this) : options.inverse(this);
                case '<':
                    return (v1 < v2) ? options.fn(this) : options.inverse(this);
                case '<=':
                    return (v1 <= v2) ? options.fn(this) : options.inverse(this);
                case '>':
                    return (v1 > v2) ? options.fn(this) : options.inverse(this);
                case '>=':
                    return (v1 >= v2) ? options.fn(this) : options.inverse(this);
                case '&&':
                    return (v1 && v2) ? options.fn(this) : options.inverse(this);
                case '||':
                    return (v1 || v2) ? options.fn(this) : options.inverse(this);
                default:
                    return options.inverse(this);
            }
        },
        debug: function(optionalValue) {
            console.log("Current Context");
            console.log("====================");
            console.log(this);

            if (optionalValue) {
                console.log("Value");
                console.log("====================");
                console.log(optionalValue);
            }
        }
    }
});

app.engine('handlebars', exbhbsEngine.engine);

app.set('view engine', 'handlebars');
app.set('views', __dirname + '/views');

app.use(express.logger());
app.use(express.bodyParser());

app.set('trust proxy', 1);
app.use(session({
    secret: 'something crazy',
    resave: false,
    saveUninitialized: true,
    cookie: { secure: false }
}));

app.use(express.static(__dirname + '/assets'));

function getXeroClient(session) {
    try {
        metaConfig = require('./config/config.json');
    } catch (ex) {
        if (process && process.env && process.env.APPTYPE) {
            //no config file found, so check the process.env.
            metaConfig.APPTYPE = process.env.APPTYPE;
            metaConfig[metaConfig.APPTYPE.toLowerCase()] = {
                authorizeCallbackUrl: process.env.authorizeCallbackUrl,
                userAgent: process.env.userAgent,
                consumerKey: process.env.consumerKey,
                consumerSecret: process.env.consumerSecret
            }
        } else {
            throw "Config not found";
        }
    }
    
    var APPTYPE = metaConfig.APPTYPE;
    var config = metaConfig[APPTYPE.toLowerCase()];
    console.log(config);
    if (session && session.token) {
        config.accessToken = session.token.oauth_token;
        config.accessSecret = session.token.oauth_token_secret;
    }

    if (config.privateKeyPath && !config.privateKey) {
        try {
            //Try to read from the path
            config.privateKey = fs.readFileSync(config.privateKeyPath);
        } catch (ex) {
            //It's not a path, so use the consumer secret as the private key
            config.privateKey = "";
        }
    }

    switch (APPTYPE) {
        case "PUBLIC":
            xeroClient = new xero.PublicApplication(config);
            break;
        case "PARTNER":
            xeroClient = new xero.PartnerApplication(config);
            eventReceiver = xeroClient.eventEmitter;
            eventReceiver.on('xeroTokenUpdate', function(data) {
                //Store the data that was received from the xeroTokenRefresh event
                console.log("Received xero token refresh: ", data);
            });
            break;
        default:
            throw "No App Type Set!!"
    }
    return xeroClient;
}

function authorizeRedirect(req, res, returnTo) {
    var xeroClient = getXeroClient(req.session, returnTo);
    xeroClient.getRequestToken(function(err, token, secret) {
        if (!err) {
            req.session.oauthRequestToken = token;
            req.session.oauthRequestSecret = secret;
            req.session.returnto = returnTo;

            //Note: only include this scope if payroll is required for your application.
            var PayrollScope = 'payroll.employees,payroll.payitems,payroll.timesheets';
            var AccountingScope = '';

            var authoriseUrl = xeroClient.buildAuthorizeUrl(token, {
                scope: AccountingScope
            });
            res.redirect(authoriseUrl);
        } else {
            res.redirect('/error');
        }
    })
}

function authorizedOperation(req, res, returnTo, callback) {
    if (req.session.token) {
      callback(getXeroClient(req.session));
    } else {
      authorizeRedirect(req, res, returnTo);
    }
}

function handleErr(err, req, res, returnTo) {
    console.log(err);
    if (err.data && err.data.oauth_problem && err.data.oauth_problem == "token_rejected") {
        authorizeRedirect(req, res, returnTo);
    } else {
        res.redirect('error', err);
    }
}

app.get('/error', function(req, res) {
    console.log(req.query.error);
    res.render('index', { error: req.query.error });
})

// Home Page
app.get('/', function(req, res) {
    res.render('index', {
        active: {
            overview: true
        }
    });
});

function getDateFormat(date){
    var dateObj = new Date(date);
    return dateFormat(dateObj, 'dd mmm yyyy');
}

app.get('/weCloudApp',function(req, res){
    console.log("Testing");
    res.end('');
})

app.get('/mapping', function(req, res){
    authorizedOperation(req, res, '/mapping', function(xeroClient) {
        //INVOICE NUMBER: 17-00001  DN-1600001 000102
        xeroClient.core.invoices.getInvoice('17-00001')
        .then(function(invoice) {
            xeroClient.core.taxRates.getTaxRates()
            .then(function(taxrates){
                //We've got the invoice so do something useful
                //Load the docx file as a binary
                var content = fs
                .readFileSync(path.resolve(__dirname, 'views/docs/weTaxInvoice.docx'), 'binary');
                console.log(content);

                var zip = new JSZip(content); 

                var doc = new Docxtemplater();
                doc.loadZip(zip);

                //Format Number of Line Items
                var lineItems = invoice.LineItems.toArray();
                var taxItems = [];
                var taxName = "";
                lineItems.forEach(function(element){
                    element.Quantity = numeral(element.Quantity).format('0,0.00');
                    element.UnitAmount = numeral(element.UnitAmount).format('0,0.00');
                    element.LineAmount = numeral(element.LineAmount).format('0,0.00');
                    taxrates.forEach(ele => {
                        if(ele.TaxType ===  element.TaxType && taxName !== ele.TaxComponents[0].Name){
                            taxName = ele.TaxComponents[0].Name;
                            let item = {
                                "TaxCode" : ele.TaxComponents[0].Name + " " + ele.TaxComponents[0].Rate + '%',
                                "TaxTotal" : numeral(element.TaxAmount).format('0,0.00')
                            }
                            taxItems.push(item);
                        }
                    });
                });
                //Get Postal Address
                var postalAddress = '';
                invoice.Contact.Addresses.forEach(element => {
                    if(element.AddressType === 'POBOX'){
                        if(typeof element.AttentionTo !== 'undefined'){
                            postalAddress += 'Attention: ' + element.AttentionTo + '\n';
                        }
                        if(typeof element.AddressLine1 !== 'undefined'){
                            postalAddress += element.AddressLine1 + '\n';
                        }
                        if(typeof element.AddressLine2 !== 'undefined'){
                            postalAddress += element.AddressLine2 + '\n';
                        }
                        if(typeof element.AddressLine3 !== 'undefined'){
                            postalAddress += element.AddressLine3 + '\n';
                        }
                        if(typeof element.AddressLine4 !== 'undefined'){
                            postalAddress += element.AddressLine4 + '\n';
                        }
                        if(typeof element.AddressLine5 !== 'undefined'){
                            postalAddress += element.AddressLine5 + '\n';
                        }
                    }
                });

                //set the templateVariables
                let data = {
                    InvoiceNumber: invoice.InvoiceNumber,
                    InvoiceDate: getDateFormat(invoice.Date),
                    ContactPostalAddress : postalAddress,
                    Reference : invoice.Reference,
                    InvoiceDueDate : getDateFormat(invoice.DueDate),
                    ContactTaxNumber: invoice.Contact.TaxNumber,
                    ContactName: invoice.Contact.Name,
                    InvoiceSubTotal: numeral(invoice.SubTotal).format('0,0.00'),
                    InvoiceTotal: numeral(invoice.Total).format('0,0.00'),
                    InvoiceTotalNetPayments : numeral(invoice.AmountCredited).format('0,0.00'),
                    InvoiceAmountDue: numeral(invoice.AmountDue).format('0,0.00'),
                    TaxItems: taxItems,
                    LineItems: lineItems
                };
                doc.setData(data);

                try {
                // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
                doc.render()
                }
                catch (error) {
                var e = {
                    message: error.message,
                    name: error.name,
                    stack: error.stack,
                    properties: error.properties,
                }
                console.log(JSON.stringify({error: e}));
                // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
                throw error;
                }

                var buf = doc.getZip()
                        .generate({type: 'nodebuffer'});

                // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
                fs.writeFileSync(path.resolve(__dirname, 'views/docs/OUTPUT.docx'), buf);
                /* var fstream = require('fs'),
                cloudconvert = new (require('cloudconvert'))('f8HO2bZ-kXDeTnDkTuzvzoMn1mzSgmStc_RhdVhFyKSxeOoHMWAtfdMs7i7zZasIUPrHy2oZO-1CrI83kQui7w');
             
                fstream.createReadStream(__dirname + '/views/docs/OUTPUT.docx')
                .pipe(cloudconvert.convert({
                    "inputformat": "docx",
                    "outputformat" : "pdf"
                }))
                .pipe(fstream.createWriteStream(__dirname + '/views/docs/OUTPUT.pdf')); */
                /* msopdf(null, function(error, office) { 

                    if (error) {
                      console.log("Init failed", error);
                      return;
                    }
                   office.word({input: __dirname + '/views/docs/OUTPUT.docx', output: __dirname + '/views/docs/OUTPUT.pdf'}, function(error, pdf) {
                      if (error) {
                           console.log("Woops", error);
                       } else {
                           console.log("Saved to", pdf);
                       }
                   });
                   office.close(null, function(error) {
                       if (error) {
                           console.log("Woops", error);
                       } else {
                           console.log("Finished & closed");
                       }
                   });
                }); */
                //res.download(__dirname + '/views/docs/OUTPUT.pdf');
            });
        });
    });
});

// Redirected from xero with oauth results
app.get('/access', function(req, res) {
    var xeroClient = getXeroClient();

    if (req.query.oauth_verifier && req.query.oauth_token == req.session.oauthRequestToken) {
        xeroClient.setAccessToken(req.session.oauthRequestToken, req.session.oauthRequestSecret, req.query.oauth_verifier)
            .then(function(token) {
                req.session.token = token.results;
                console.log(req.session);

                var returnTo = req.session.returnto;
                res.redirect(returnTo || '/');
            })
            .catch(function(err) {
                handleErr(err, req, res, 'error');
            })
    }
});

app.get('/organisations', function(req, res) {
    authorizedOperation(req, res, '/organisations', function(xeroClient) {
        xeroClient.core.organisations.getOrganisations()
            .then(function(organisations) {
                res.render('organisations', {
                    organisations: organisations,
                    active: {
                        organisations: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'organisations');
            })
    })
});

app.get('/brandingthemes', function(req, res) {
    authorizedOperation(req, res, '/brandingthemes', function(xeroClient) {
        xeroClient.core.brandingThemes.getBrandingThemes()
            .then(function(brandingthemes) {
                res.render('brandingthemes', {
                    brandingthemes: brandingthemes,
                    active: {
                        brandingthemes: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'brandingthemes');
            })
    })
});

app.get('/invoicereminders', function(req, res) {
    authorizedOperation(req, res, '/invoicereminders', function(xeroClient) {
        xeroClient.core.invoiceReminders.getInvoiceReminders()
            .then(function(invoiceReminders) {
                res.render('invoicereminders', {
                    invoicereminders: invoiceReminders,
                    active: {
                        invoicereminders: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'invoicereminders');
            })
    })
});

app.get('/taxrates', function(req, res) {
    authorizedOperation(req, res, '/taxrates', function(xeroClient) {
        xeroClient.core.taxRates.getTaxRates()
            .then(function(taxrates) {
                res.render('taxrates', {
                    taxrates: taxrates,
                    active: {
                        taxrates: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'taxrates');
            })
    })
});

app.get('/users', function(req, res) {
    authorizedOperation(req, res, '/users', function(xeroClient) {
        xeroClient.core.users.getUsers()
            .then(function(users) {
                res.render('users', {
                    users: users,
                    active: {
                        users: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'users');
            })
    })
});

app.get('/contacts', function(req, res) {
    authorizedOperation(req, res, '/contacts', function(xeroClient) {
        var contacts = [];
        xeroClient.core.contacts.getContacts({ pager: { callback: pagerCallback } })
            .then(function() {
                res.render('contacts', {
                    contacts: contacts,
                    active: {
                        contacts: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'contacts');
            })

        function pagerCallback(err, response, cb) {
            contacts.push.apply(contacts, response.data);
            cb()
        }
    })
});

app.get('/currencies', function(req, res) {
    authorizedOperation(req, res, '/currencies', function(xeroClient) {
        xeroClient.core.currencies.getCurrencies()
            .then(function(currencies) {
                res.render('currencies', {
                    currencies: currencies,
                    active: {
                        currencies: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'currencies');
            });
    })
});

app.get('/banktransactions', function(req, res) {
    authorizedOperation(req, res, '/banktransactions', function(xeroClient) {
        var bankTransactions = [];
        xeroClient.core.bankTransactions.getBankTransactions({ pager: { callback: pagerCallback } })
            .then(function() {
                res.render('banktransactions', {
                    bankTransactions: bankTransactions,
                    active: {
                        banktransactions: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'banktransactions');
            })

        function pagerCallback(err, response, cb) {
            bankTransactions.push.apply(bankTransactions, response.data);
            cb()
        }
    })
});

app.get('/journals', function(req, res) {
    authorizedOperation(req, res, '/journals', function(xeroClient) {
        var journals = [];
        xeroClient.core.journals.getJournals({ pager: { callback: pagerCallback } })
            .then(function() {
                res.render('journals', {
                    journals: journals,
                    active: {
                        journals: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'journals');
            })

        function pagerCallback(err, response, cb) {
            journals.push.apply(journals, response.data);
            cb()
        }
    })
});

app.get('/banktransfers', function(req, res) {
    authorizedOperation(req, res, '/banktransfers', function(xeroClient) {
        var bankTransfers = [];
        xeroClient.core.bankTransfers.getBankTransfers({ pager: { callback: pagerCallback } })
            .then(function() {
                res.render('banktransfers', {
                    bankTransfers: bankTransfers,
                    active: {
                        banktransfers: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'banktransfers');
            })

        function pagerCallback(err, response, cb) {
            bankTransfers.push.apply(bankTransfers, response.data);
            cb()
        }
    })
});

app.get('/payments', function(req, res) {
    authorizedOperation(req, res, '/payments', function(xeroClient) {
        xeroClient.core.payments.getPayments()
            .then(function(payments) {
                res.render('payments', {
                    payments: payments,
                    active: {
                        payments: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'payments');
            })
    })
});

app.get('/trackingcategories', function(req, res) {
    authorizedOperation(req, res, '/trackingcategories', function(xeroClient) {
        xeroClient.core.trackingCategories.getTrackingCategories()
            .then(function(trackingcategories) {
                res.render('trackingcategories', {
                    trackingcategories: trackingcategories,
                    active: {
                        trackingcategories: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'trackingcategories');
            })
    })
});

app.get('/accounts', function(req, res) {
    authorizedOperation(req, res, '/accounts', function(xeroClient) {
        xeroClient.core.accounts.getAccounts()
            .then(function(accounts) {
                res.render('accounts', {
                    accounts: accounts,
                    active: {
                        accounts: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'accounts');
            })
    })
});

app.get('/creditnotes', function(req, res) {
    authorizedOperation(req, res, '/creditnotes', function(xeroClient) {
        xeroClient.core.creditNotes.getCreditNotes()
            .then(function(creditnotes) {
                res.render('creditnotes', {
                    creditnotes: creditnotes,
                    active: {
                        creditnotes: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'creditnotes');
            })
    })
});

app.get('/invoices', function(req, res) {
    authorizedOperation(req, res, '/invoices', function(xeroClient) {
        xeroClient.core.invoices.getInvoice('17-00001')
        .then(function(inv) {
            //We've got the invoice so do something useful
            //console.log(invoice); //ACCPAY
            /* xeroClient.core.contacts.getContact('e73a1e98-d87e-4075-8290-a852dfadcc59')
            .then(function(contact){
                if(typeof contact.Addresses[1].AttentionTo !== 'undefined'){
                    res.json(contact.Addresses[1].AttentionTo);
                }else{
                    res.end('xxx');
                }
            }); */
            res.json(inv);
        });
        /* xeroClient.core.invoices.getInvoices()
            .then(function(invoices) {
                res.render('invoices', {
                    invoices: invoices,
                    active: {
                        invoices: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'invoices');
            }) */

    })
});

app.get('/repeatinginvoices', function(req, res) {
    authorizedOperation(req, res, '/repeatinginvoices', function(xeroClient) {
        xeroClient.core.repeatinginvoices.getRepeatingInvoices()
            .then(function(repeatingInvoices) {
                res.render('repeatinginvoices', {
                    repeatinginvoices: repeatingInvoices,
                    active: {
                        repeatinginvoices: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'repeatinginvoices');
            })

    })
});

app.get('/attachments', function(req, res) {
    authorizedOperation(req, res, '/attachments', function(xeroClient) {

        var entityID = req.query && req.query.entityID ? req.query.entityID : null;
        var entityType = req.query && req.query.entityType ? req.query.entityType : null;

        if (entityID && entityType) {

            xeroClient.core.invoices.getInvoice(entityID)
                .then(function(invoice) {
                    invoice.getAttachments()
                        .then(function(attachments) {
                            res.render('attachments', {
                              attachments: attachments,
                              InvoiceID: entityID,
                              active: {
                                  invoices: true,
                                  nav: {
                                      accounting: true
                                  }
                              }
                          });
                        })
                        .catch(function(err) {
                            handleErr(err, req, res, 'attachments');
                        })
                })
                .catch(function(err) {
                    handleErr(err, req, res, 'attachments');
                })

        } else {
            handleErr("No Attachments Found", req, res, 'index');
        }
    })
});

app.get('/download', function(req, res) {
    authorizedOperation(req, res, '/attachments', function(xeroClient) {

        var entityID = req.query && req.query.entityID ? req.query.entityID : null;
        var entityType = req.query && req.query.entityType ? req.query.entityType : null;
        var fileId = req.query && req.query.fileId ? req.query.fileId : null;

        if (entityID && entityType && fileId) {

            xeroClient.core.invoices.getInvoice(entityID)
                .then(function(invoice) {
                    invoice.getAttachments()
                        .then(function(attachments) {
                            attachments.forEach(attachment => {
                              //Get the reference to the attachment object
                              if(attachment.AttachmentID === fileId) {
                                res.writeHead(200, {
                                    "Content-Type": attachment.MimeType,
                                    "Content-Disposition": "attachment; filename=" + attachment.FileName,
                                    "Content-Length": attachment.ContentLength
                                });
                                attachment.getContent(res);
                              }
                            });
                        })
                        .catch(function(err) {
                            handleErr(err, req, res, 'attachments');
                        })
                })
                .catch(function(err) {
                    handleErr(err, req, res, 'attachments');
                })

        } else {
            handleErr("No Attachments Found", req, res, 'index');
        }
    })
});

app.get('/items', function(req, res) {
    authorizedOperation(req, res, '/items', function(xeroClient) {
        xeroClient.core.items.getItems()
            .then(function(items) {
                res.render('items', {
                    items: items,
                    active: {
                        items: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'items');
            })

    })
});

app.get('/manualjournals', function(req, res) {
    authorizedOperation(req, res, '/manualjournals', function(xeroClient) {
        var manualjournals = [];
        xeroClient.core.manualjournals.getManualJournals({ pager: { callback: pagerCallback } })
            .then(function() {
                res.render('manualjournals', {
                    manualjournals: manualjournals,
                    active: {
                        manualjournals: true,
                        nav: {
                            accounting: true
                        }
                    }
                });
            })
            .catch(function(err) {
                handleErr(err, req, res, 'manualjournals');
            })

        function pagerCallback(err, response, cb) {
            manualjournals.push.apply(manualjournals, response.data);
            cb()
        }
    })
});

app.get('/reports', function(req, res) {
    authorizedOperation(req, res, '/reports', function(xeroClient) {

        var reportkeys = {
            '1': 'BalanceSheet',
            '2': 'TrialBalance',
            '3': 'ProfitAndLoss',
            '4': 'BankStatement',
            '5': 'BudgetSummary',
            '6': 'ExecutiveSummary',
            '7': 'BankSummary',
            '8': 'AgedReceivablesByContact',
            '9': 'AgedPayablesByContact',
            '10': 'TenNinetyNine'
        };

        var report = req.query ? req.query.r : null;

        if (reportkeys[report]) {
            var selectedReport = reportkeys[report];

            var data = {
                active: {
                    nav: {
                        reports: true
                    }
                }
            };

            data.active[selectedReport.toLowerCase()] = true;

            /**
             * We may need some dependent data:
             * 
             * BankStatement - requires a BankAccountId
             * AgedReceivablesByContact - requires a ContactId
             * AgedPayablesByContact - requires a ContactId
             * 
             */

            if (selectedReport == 'BankStatement') {
                xeroClient.core.accounts.getAccounts({ where: 'Type=="BANK"' })
                    .then(function(accounts) {
                        xeroClient.core.reports.generateReport({
                                id: selectedReport,
                                params: {
                                    bankAccountID: accounts[0].AccountID
                                }
                            })
                            .then(function(report) {
                                data.report = report.toObject();
                                data.colspan = data.report.Rows[0].Cells.length;
                                res.render('reports', data);
                            })
                            .catch(function(err) {
                                handleErr(err, req, res, 'reports');
                            });
                    })
                    .catch(function(err) {
                        handleErr(err, req, res, 'reports');
                    });
            } else if (selectedReport == 'AgedReceivablesByContact' || selectedReport == 'AgedPayablesByContact') {
                xeroClient.core.contacts.getContacts()
                    .then(function(contacts) {
                        xeroClient.core.reports.generateReport({
                                id: selectedReport,
                                params: {
                                    contactID: contacts[0].ContactID
                                }
                            })
                            .then(function(report) {
                                data.report = report.toObject();
                                data.colspan = data.report.Rows[0].Cells.length;
                                res.render('reports', data);
                            })
                            .catch(function(err) {
                                handleErr(err, req, res, 'reports');
                            });
                    })
                    .catch(function(err) {
                        handleErr(err, req, res, 'reports');
                    });
            } else {
                xeroClient.core.reports.generateReport({
                        id: selectedReport
                    })
                    .then(function(report) {
                        data.report = report.toObject();
                        if (data.report.Rows) {
                            data.colspan = data.report.Rows[0].Cells.length;
                        }
                        res.render('reports', data);
                    })
                    .catch(function(err) {
                        handleErr(err, req, res, 'reports');
                    });
            }

        } else {
            res.render('index', {
                error: {
                    message: "Report not found"
                },
                active: {
                    overview: true
                }
            });
        }
    })
});

app.use('/createinvoice', function(req, res) {
    if (req.method == 'GET') {
        return res.render('createinvoice');
    } else if (req.method == 'POST') {
        authorizedOperation(req, res, '/createinvoice', function(xeroClient) {
            var invoice = xeroClient.core.invoices.newInvoice({
                Type: req.body.Type,
                Contact: {
                    Name: req.body.Contact
                },
                DueDate: '2014-10-01',
                LineItems: [{
                    Description: req.body.Description,
                    Quantity: req.body.Quantity,
                    UnitAmount: req.body.Amount,
                    AccountCode: 400,
                    ItemCode: 'ABC123'
                }],
                Status: 'DRAFT'
            });
            invoice.save()
                .then(function(ret) {
                    res.render('createinvoice', { outcome: 'Invoice created', id: ret.entities[0].InvoiceID })
                })
                .catch(function(err) {
                    res.render('createinvoice', { outcome: 'Error', err: err })
                })

        })
    }
});

app.use(function(req, res, next) {
    if (req.session)
        delete req.session.returnto;
})

var PORT = process.env.PORT || 3100;

app.listen(PORT);
console.log("listening on http://localhost:" + PORT);
