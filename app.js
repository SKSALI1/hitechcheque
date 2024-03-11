const express = require("express");
const ExcelJS = require("exceljs");
const bodyParser = require("body-parser");
const ejs = require("ejs");
const mongoose = require("mongoose");
const flash = require("connect-flash");
const session = require("express-session");


const app = express();



app.use(
  session({
    secret: "hgfsuadfsdufklasdufasd",
    saveUninitialized: true,
    resave: true,
  })
);
app.use(flash());
app.use((req, res, next) => {
  (res.locals.success = req.flash("success")),
    (res.locals.error = req.flash("err"));
  next();
});
app.use(express.static(__dirname + '/public'));
app.use('/edit-entry', express.static('public/edit-entry'));




app.set("view engine", "ejs");
app.use(
  bodyParser.urlencoded({
    extended: true,
  })
);
// mongoose.connect("mongodb://localhost:27017/chequeDB", {
//  useNewUrlParser: true,
// });

 mongoose.connect("mongodb+srv://admin-saahil:Test123@cluster0.r5ume9x.mongodb.net/chequeDB", {
   useNewUrlParser: true,
 });
const entryNumberSchema = new mongoose.Schema({
  currentEntryNumber: Number,
});

const EntryNumber = mongoose.model("EntryNumber", entryNumberSchema);

const chequeSchema = {
  date: String,
  entryId: Number,
  acnum: String,
  acname: String,
  amount: String,
  amountinwords: String,
  totalamount: String,
  unisolvenumber: String,

  firstcompanyname: String,
  firstbankname: String,
  firstpartyname: String,
  firstbranchname: String,
  firstchequenumber: String,
  firstchequedate: String,
  firstamount: Number,

  iicompanyname: String,
  secondpartyname: String,
  secondbankname: String,
  secondbranchname: String,
  secondchequenumber: String,
  secondchequedate: String,
  secondamount: Number,

  iiicompanyname: String,
  thirdpartyname: String,
  thirdbankname: String,
  thirdbranchname: String,
  thirdchequenumber: String,
  thirdchequedate: String,
  thirdamount: Number,

  ivcompanyname: String,
  fourthpartyname: String,
  fourthbankname: String,
  fourthbranchname: String,
  fourthchequenumber: String,
  fourthchequedate: String,
  fourthamount: Number,

  vcompanyname: String,
  fifthpartyname: String,
  fifthbankname: String,
  fifthbranchname: String,
  fifthchequenumber: String,
  fifthchequedate: String,
  fifthamount: Number,
};
const Chequeentry = new mongoose.model("Chequeentry", chequeSchema);

const customerSchema = new mongoose.Schema({
  partyname:String,
  partyarea:String
})

const Customer = mongoose.model("Customer", customerSchema);

// app.get("/", function (req, res) {
//   req.flash('message','success message')
//   res.redirect("home");
// });
app.get("/", function (req, res) {
  res.render("home");
});
app.get('/search', (req, res) => {
  res.render('search'); // Renders the search.ejs file
});

app.get('/search-data', async (req, res) => {
  const searchTerm = req.query.query;

  try {
      const parties = await Customer.find({ partyname: { $regex: searchTerm, $options: 'i' } }, 'partyname');

      const partyNames = parties.map(party => party.partyname);

      res.json(partyNames);
  } catch (error) {
      console.error('Error fetching party names:', error);
      res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.get("/edit-entry", function (req, res) {
  res.render("edit-entry");
});

app.get("/register", async (req, res) => {
  try {
    const entryNumberDoc = await EntryNumber.findOne();
    const currentEntryNumber = entryNumberDoc
      ? entryNumberDoc.currentEntryNumber
      : 1;
    res.render("register", { currentEntryNumber: currentEntryNumber });
  } catch (err) {
    console.error(err);
  }
});

app.get("/details", function (req, res) {
  // Provide a default entry object here
  const entry = {
    entryId: "",
    date: "",
    acnum: "",
    acname: "",
    amount: "",
    amountinwords: "",
    unisolvenumber: "",
    firstbankname: "",
    firstbranchname: "",
    firstchequenumber: "",
    firstamount: "",
    secondbankname: "",
    secondbranchname: "",
    secondchequenumber: "",
    secondamount: "",
    thirdbankname: "",
    thirdbranchname: "",
    thirdchequenumber: "",
    thirdamount: "",
    fourthbankname: "",
    fourthbranchname: "",
    fourthchequenumber: "",
    fourthamount: "",
    fifthbankname: "",
    fifthbranchname: "",
    fifthchequenumber: "",
    fifthamount: "",
  };
  res.render("details", { entry: entry });
});

app.get("/getdetails", async (req, res) => {
  try {
    const entryId = req.query.entryId;
    let errorMessage = null;

    if (entryId) {
      const entry = await Chequeentry.findOne({ entryId: parseInt(entryId) });
      if (entry) {
        res.render("details", { entry: entry, errorMessage: errorMessage });
      } else {
        errorMessage = "Entry not found";
        res.render("details", { entry: null, errorMessage: errorMessage });
      }
    }
  } catch (err) {
    console.error(err);
    errorMessage = "An error occurred";
    res.render("details", { entry: null, errorMessage: errorMessage });
  }
});

app.get("/details/:id", async (req, res) => {
  try {
    const entry = await Chequeentry.findOne({ entryId: req.params.id });
    if (entry) {
      res.render("details", { entry: entry });
    } else {
      res.send("entry not foumd")
    }
  } catch (err) {
    console.error(err);
  }
});

app.get("/allentries", async (req, res) => {
  try {
    let filter = {};
    if (req.query.unisolvenumber) {
      filter.unisolvenumber = req.query.unisolvenumber;
    }
    if (req.query.entryId) {
      filter.entryId = req.query.entryId;
    }
    if (req.query.acname) {
      filter.acname = req.query.acname;
    }

    // Check if startDate and endDate are provided in the query
    if (req.query.startDate && req.query.endDate) {
      filter.date = {
        $gte: req.query.startDate,
        $lte: req.query.endDate,
      };
    }

    const entries = await Chequeentry.find(filter)
      .sort({ entryId: -1 })
      .maxTimeMS(300000);
    res.render("allentries", { entries: entries });
  } catch (err) {
    console.error(err);
    // Handle errors as needed
  }
});

app.get("/download-excel", async (req, res) => {
  try {
    const entries = await Chequeentry.find();

    const ExcelJS = require("exceljs");
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("All Entries");

    worksheet.addRow([
      "Entry ID",
      "Date",
      "Account Number",
      "Account Name",
      "totalAmount",
      "unisolvenumber",
      "CHQ1COMPANYNAME",
      "CHQ1BANKNAME",
      "CHQ1PARTYNAME",
      "CHQ1BRANCHNAME",
      "CHQ2COMPANYNAME",
      "CHQ2BANKNAME",
      "CHQ2PARTYNAME",
      "CHQ2BRANCHNAME",
      "CHQ3COMPANYNAME",
      "CHQ3BANKNAME",
      "CHQ3PARTYNAME",
      "CHQ3BRANCHNAME",
      "CHQ4COMPANYNAME",
      "CHQ4BANKNAME",
      "CHQ4PARTYNAME",
      "CHQ4BRANCHNAME"
    ]); // Add column headers

    entries.forEach((entry) => {
      worksheet.addRow([
        entry.entryId,
        entry.date,
        entry.acnum,
        entry.acname,
        entry.totalAmount,
        entry.unisolvenumber,
        entry.firstcompanyname,
        entry.firstbankname,
        entry.firstpartyname,
        entry.firstbranchname,
        entry.iicompanyname,
        entry.secondbankname,
        entry.secondpartyname,
        entry.secondbranchname,
        entry.iiicompanyname,
        entry.thirdbankname,
        entry.thirdpartyname,
        entry.thirdbranchname,
        entry.ivcompanyname,
        entry.fourthbankname,
        entry.fourthpartyname,
        entry.fourthbranchname,
      ]); // Add data rows
    });

    // Set headers for the download
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=all-entries.xlsx"
    );

    // Write the Excel file to the response
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    // Handle errors as needed
    res.status(500).send("An error occurred");
  }
});

app.get("/edit-entry/:id", async (req, res) => {
  try {
    const entry = await Chequeentry.findById(req.params.id);
    res.render("edit-entry", { entry }); 
  } catch (error) {
    console.error(error);
    res.status(500).send("Internal Server Error");
  }
});

app.get("/addparty", function (req, res){
  res.render("addparty");
});

app.get("/excel", async (req, res) => {
  try {
    let filter = {}; // Check if startDate and endDate are provided in the query
    if (req.query.startDate && req.query.endDate) {
      filter.date = {
        $gte: req.query.startDate,
        $lte: req.query.endDate,
      };
    }
    const entries = await Chequeentry.find(filter)
      .sort({ entryId: -1 })
      .maxTimeMS(300000);
    res.render("excel", { entries: entries });
  } catch (err) {
    console.error(err);
    // Handle errors as needed
  }
});








app.post("/register", async (req, res) => {
  try {
    let entryNumberDoc = await EntryNumber.findOne();

    if (!entryNumberDoc) {
      entryNumberDoc = new EntryNumber({
        currentEntryNumber: 1,
      });
    }

    const newEntryNumber = entryNumberDoc.currentEntryNumber;
    entryNumberDoc.currentEntryNumber++;
    const firstAmount = parseFloat(req.body.firstamount) || 0;
    const secondAmount = parseFloat(req.body.secondamount) || 0;
    const thirdAmount = parseFloat(req.body.thirdamount) || 0;
    const fourthAmount = parseFloat(req.body.fourthamount) || 0;
    const fifthAmount = parseFloat(req.body.fifthamount) || 0;
    const totalAmount = firstAmount + secondAmount + thirdAmount + fourthAmount + fifthAmount;

    const a = [
      "",
      "ONE ",
      "TWO ",
      "THREE ",
      "FOUR ",
      "FIVE ",
      "SIX ",
      "SEVEN ",
      "EIGHT ",
      "NINE ",
      "TEN ",
      "ELEVEN ",
      "TWELVE ",
      "THIRTEEN ",
      "FOURTEEN ",
      "FIFTEEN ",
      "SIXTEEN ",
      "SEVENTEEN ",
      "EIGHTEEN ",
      "NINETEEN ",
    ];
    const b = [
      "",
      "",
      "TWENTY",
      "THIRTY",
      "FORTY",
      "FIFTY",
      "SIXTY",
      "SEVENTY",
      "EIGHTY",
      "NINETY",
    ];

    function inWords(num) {
      if ((num = num.toString()).length > 9) return "overflow";
      n = ("000000000" + num)
        .substr(-9)
        .match(/^(\d{2})(\d{2})(\d{2})(\d{1})(\d{2})$/);
      if (!n) return;
      var str = "";
      str +=
        n[1] != 0
          ? (a[Number(n[1])] || b[n[1][0]] + " " + a[n[1][1]]) + "CRORE "
          : "";
      str +=
        n[2] != 0
          ? (a[Number(n[2])] || b[n[2][0]] + " " + a[n[2][1]]) + "LAKH "
          : "";
      str +=
        n[3] != 0
          ? (a[Number(n[3])] || b[n[3][0]] + " " + a[n[3][1]]) + "THOUSAND "
          : "";
      str +=
        n[4] != 0
          ? (a[Number(n[4])] || b[n[4][0]] + " " + a[n[4][1]]) + "HUNDRED "
          : "";
      str +=
        n[5] != 0
          ? (str != "" ? "& " : "") +
            (a[Number(n[5])] || b[n[5][0]] + " " + a[n[5][1]]) +
            "ONLY "
          : "";
      return str;
    }

    const totalAmountInWords = inWords(totalAmount);

    await entryNumberDoc.save();

    const newUser = new Chequeentry({
      date: req.body.date,
      entryId: newEntryNumber,
      acnum: req.body.acnum,
      acname: req.body.acname ? req.body.acname.toUpperCase() : "",
      amount: req.body.amount,
      amountinwords: totalAmountInWords,
      unisolvenumber: "",

      firstcompanyname: req.body.firstcompanyname
        ? req.body.firstcompanyname.toUpperCase()
        : " ",
      firstpartyname: req.body.firstpartyname,
      firstbankname: req.body.firstbankname,
      firstbranchname: req.body.firstbranchname,
      firstchequenumber: req.body.firstchequenumber,
      firstchequedate: req.body.firstchequedate,
      firstamount: req.body.firstamount,

      iicompanyname: req.body.iicompanyname,
      secondpartyname: req.body.secondpartyname,
      secondbankname: req.body.secondbankname,
      secondbranchname: req.body.secondbranchname,
      secondchequenumber: req.body.secondchequenumber,
      secondchequedate: req.body.secondchequedate,
      secondamount: req.body.secondamount,

      iiicompanyname: req.body.iiicompanyname,
      thirdpartyname: req.body.thirdpartyname,
      thirdbankname: req.body.thirdbankname,
      thirdbranchname: req.body.thirdbranchname,
      thirdchequenumber: req.body.thirdchequenumber,
      thirdchequedate: req.body.thirdchequedate,
      thirdamount: req.body.thirdamount,

      ivcompanyname: req.body.ivcompanyname,
      fourthpartyname: req.body.fourthpartyname,
      fourthbankname: req.body.fourthbankname,
      fourthbranchname: req.body.fourthbranchname,
      fourthchequenumber: req.body.fourthchequenumber,
      fourthchequedate: req.body.fourthchequedate,
      fourthamount: req.body.fourthamount,

      vcompanyname: req.body.vcompanyname,
      fifthpartyname: req.body.fifthpartyname,
      fifthbankname: req.body.fifthbankname,
      fifthbranchname: req.body.fifthbranchname,
      fifthchequenumber: req.body.fifthchequenumber,
      fifthchequedate: req.body.fifthchequedate,
      fifthamount: req.body.fifthamount,

      // Store the total amount in the database
      totalamount: totalAmount,
    });

    await newUser.save();
    req.flash("success", "Entry Added Successfully"),
      res.redirect("/allentries");
  } catch (err) {
    req.flash("err", "Entry Added Successfully"), res.redirect("/allentries");
    console.error(err);
  }
});

app.post("/allentries", async (req, res) => {
  try {
    const entryId = req.body.entryId;
    const newUnisolveNumber = req.body.unisolvenumber;
    const entry = await Chequeentry.findOne({ entryId: entryId });

    if (entry) {
      entry.unisolvenumber = newUnisolveNumber;
      await entry.save();
      // req.flash("success", "Unisolvenumber updated successfully!");
      // res.send(req.flash('message'));
      res.status(204).send();
    } else {
      res.status(404).send("Entry not found");
    }
  } catch (err) {
    console.error(err);
    res.status(500).send("An error occurred");
  }
});

app.post("/deleteentry/:entryId", async (req, res) => {
  try {
    const entryId = req.params.entryId;
    const deletedEntry = await Chequeentry.findByIdAndRemove(entryId);

    if (!deletedEntry) {
      res.status(404).send("Entry not found");
    } else {
      res.redirect("/allentries");
    }
  } catch (err) {
    console.error(err);
    // Handle errors as needed
    res.status(500).send("Internal Server Error");
  }
});


app.post("/update-entry/:id", async (req, res) => {
  try {
    const updatedEntry = await Chequeentry.findByIdAndUpdate(
      req.params.id,
      req.body,
      { new: true }
    );

    // Recalculate the total 
    const firstAmount = parseFloat(req.body.firstamount) || 0;
    const secondAmount = parseFloat(req.body.secondamount) || 0;
    const thirdAmount = parseFloat(req.body.thirdamount) || 0;
    const fourthAmount = parseFloat(req.body.fourthamount) || 0;
    const fifthAmount = parseFloat(req.body.fifthamount) || 0;

    // Update the total 
    updatedEntry.totalamount = firstAmount + secondAmount + thirdAmount + fourthAmount + fifthAmount;

    // Save the updated entry
    await updatedEntry.save();

    res.redirect("/allentries"); 
  } catch (error) {
    console.error(error);
    res.status(500).send("Internal Server Error");
  }
});


app.post("/newparty", async (req, res) => {
  try {
    const newCustomer = new Customer({
      partyname: req.body.newparty,
      partyarea: req.body.newpartyarea
    });

    const savedCustomer = await newCustomer.save();

    res.send("Party added successfully.");
  } catch (error) {
    console.error('Error adding party:', error);
    res.status(500).json({ error: error.message });
  }
});





app.listen(3000, function () {
  console.log("Server started on port 3000.");
});
