const admin = require("firebase-admin");
const nodemailer = require("nodemailer");
const ExcelJS = require("exceljs");
require('dotenv').config();

// Firebase is initialized lazily in the direct-run guard at the bottom so this
// module can be imported (e.g. by tests) without service-account credentials.
let db;

const StatementType = "Monthly"; // Change to "Weekly" for weekly statements


function getPreviousWeekRange() {
    const now = new Date();

    // Get today's date in IST (date-only context)
    const istToday = new Date(
        now.toLocaleDateString("en-CA", { timeZone: "Asia/Kolkata" })
    );

    const day = istToday.getDay(); // 0=Sun ... 6=Sat

    // Find last Friday
    const diffToFriday = (day >= 5) ? (day - 5) : (day + 2);

    const endDate = new Date(istToday);
    endDate.setDate(istToday.getDate() - diffToFriday);

    const startDate = new Date(endDate);
    startDate.setDate(endDate.getDate() - 6);

    // Format as YYYY-MM-DD
    const start = startDate.toISOString().slice(0, 10);
    const end = endDate.toISOString().slice(0, 10);

    // const start = '2026-02-07';
    // const end = '2026-02-13';

    return { start, end };
}
function getPreviousMonthRange() {

    const now = new Date();

    // Get today's date in IST
    const istToday = new Date(
        now.toLocaleDateString("en-CA", { timeZone: "Asia/Kolkata" })
    );

    const year = istToday.getFullYear();
    const month = istToday.getMonth();

    // First day of previous month
    const startDate = new Date(year, month - 1, 1);

    // Last day of previous month
    const endDate = new Date(year, month, 0);

    const start = startDate.toLocaleDateString("en-CA");
    const end = endDate.toLocaleDateString("en-CA");

    // const start = '2025-12-01';
    // const end = '2025-12-31';

    return { start, end };
}
function calculateDuration(startTime, endTime) {
    if (!startTime || !endTime) return "";

    // startTime & endTime format: "HH:mm"
    const startParts = startTime.split(":");
    const endParts = endTime.split(":");

    const startMinutes = parseInt(startParts[0], 10) * 60 + parseInt(startParts[1], 10);
    const endMinutes = parseInt(endParts[0], 10) * 60 + parseInt(endParts[1], 10);

    let diffMinutes = endMinutes - startMinutes;

    // Handle overnight slots (example: 23:00 → 01:00)
    if (diffMinutes < 0) {
        diffMinutes += 24 * 60;
    }

    const hours = diffMinutes / 60;

    if (hours === 1) return "1 Hour";
    return hours + " Hours";
}
// A cancelled booking whose online payment was credited back to the customer's
// GameOn Wallet. Its money must be excluded from report revenue.
function isRefundedToWallet(b) {
    const isCancelled = (b.status || "").toLowerCase() === "cancelled";
    if (!isCancelled) return false;
    return (
        b.wallet_refund_choice === "refund_to_gameon_wallet" ||
        Number(b.wallet_refund_amount || 0) > 0
    );
}
function splitPayments(payments = []) {
    let razorpay = 0,
        cash = 0,
        upi = 0,
        card = 0,
        wallet = 0;

    payments.forEach(p => {
        const amount = Number(p.amount || 0);
        const method = (p.method || "").toLowerCase();

        // "Online" is the customer-facing label for gateway (Razorpay) card
        // payments; older records may still carry "razorpay".
        if (method.includes("razor") || method === "online") razorpay += amount;
        else if (method === "cash") cash += amount;
        else if (method === "upi") upi += amount;
        else if (method === "card") card += amount;
        // GameOn Wallet is an online payment method; keep it in its own bucket
        // so it counts toward Total Paid without being mislabelled as Razorpay.
        else if (method.includes("wallet")) wallet += amount;
    });

    return { razorpay, cash, upi, card, wallet };
}
function createBookingSheet(wb, sheetName, bookings) {
    const ws = wb.addWorksheet(sheetName);

    const header = [
        "Booking ID", "Booking Type", "Booking Date", "Booking Status",
        "Venue Name", "Court ID", "Sport", "Start Time", "End Time",
        "Duration", "Slot Cost", "Final Amount",
        "Coupon Code", "Coupon Discount",
        "Online", "Cash", "UPI", "Card", "GameOn Wallet",
        "Total Paid", "Balance/Due",
        "Customer Name", "Customer Phone", "Country Code"
    ];

    const headerRow = ws.addRow(header);
    headerRow.eachCell(cell => {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF000000" } };
        cell.alignment = { vertical: "middle", horizontal: "center" };
    });

    let totals = {
        slotCost: 0,
        finalAmount: 0,
        couponDiscount: 0,
        razorpay: 0,
        cash: 0,
        upi: 0,
        card: 0,
        wallet: 0,
        totalPaid: 0,
        balance: 0
    };

    bookings.forEach(b => {
        const payments = splitPayments(b.payment || []);

        const isCancelled = (b.status || "").toLowerCase() === "cancelled";
        // Online money refunded to the customer's GameOn Wallet is no longer
        // revenue — drop the Razorpay + Wallet portions so the report matches
        // the settlement/finance behaviour.
        if (isRefundedToWallet(b)) {
            payments.razorpay = 0;
            payments.wallet = 0;
        }
        const totalPaid = payments.razorpay + payments.cash + payments.upi + payments.card + payments.wallet;

        const couponDiscount = Number(b.coupon_discount) || 0;
        const finalAmount = isCancelled ? 0 : (Number(b.final_amount) || 0);
        const balance = isCancelled ? 0 : finalAmount - totalPaid;

        totals.slotCost += Number(b.slot_cost) || 0;
        totals.finalAmount += finalAmount;
        totals.couponDiscount += couponDiscount;
        totals.razorpay += payments.razorpay;
        totals.cash += payments.cash;
        totals.upi += payments.upi;
        totals.card += payments.card;
        totals.wallet += payments.wallet;
        totals.totalPaid += totalPaid;
        totals.balance += balance;

        const row = ws.addRow([
            b.id,
            b.booking_type,
            b.date,
            b.status,
            b.venue_name,
            b.court_id,
            (b.sport && b.sport.name) ? b.sport.name : "",
            b.start_time,
            b.end_time,
            calculateDuration(b.start_time, b.end_time),
            b.slot_cost,
            finalAmount,
            b.coupon_code,
            couponDiscount,
            payments.razorpay,
            payments.cash,
            payments.upi,
            payments.card,
            payments.wallet,
            totalPaid,
            balance,
            (b.user && b.user.name) ? b.user.name : "",
            (b.user && b.user.phone) ? b.user.phone : "",
            (b.user && b.user.country_code) ? b.user.country_code : ""
        ]);

        row.getCell(3).numFmt = "yyyy-mm-dd";
        [11, 12, 14, 15, 16, 17, 18, 19, 20, 21].forEach(i => {
            row.getCell(i).numFmt = "#,##0.00";
        });

        if (isCancelled) {
            row.eachCell(cell => {
                cell.font = { color: { argb: "FF888888" } };
            });
        }
    });

    const totalRow = ws.addRow([
        "TOTAL", "", "", "", "", "", "", "", "",
        "",
        totals.slotCost,
        totals.finalAmount,
        "",
        totals.couponDiscount,
        totals.razorpay,
        totals.cash,
        totals.upi,
        totals.card,
        totals.wallet,
        totals.totalPaid,
        totals.balance,
        "", "", ""
    ]);

    totalRow.eachCell(cell => {
        cell.font = { bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF3CD" } };
        cell.border = { top: { style: "thin" }, bottom: { style: "double" } };
    });

    [11, 12, 14, 15, 16, 17, 18, 19, 20, 21].forEach(i => {
        totalRow.getCell(i).numFmt = "#,##0.00";
    });

    ws.columns.forEach(col => {
        let maxLength = 0;
        col.eachCell({ includeEmpty: true }, cell => {
            maxLength = Math.max(maxLength, (cell.value || "").toString().length);
        });
        col.width = maxLength < 10 ? 10 : maxLength + 2;
    });
}
function createOnlineBookingSheet(wb, sheetName, bookings, commission_percentage) {
    const ws = wb.addWorksheet(sheetName);

    const header = [
        "Booking ID", "Booking Type", "Booking Date", "Booking Status",
        "Venue Name", "Court ID", "Sport", "Start Time", "End Time",
        "Duration", "Slot Cost", "Final Amount",
        "Coupon Code", "Coupon Discount",
        "Online", "Cash", "UPI", "Card", "GameOn Wallet",
        "Total Paid", "Balance/Due", "Commission Percentage", "Commission Amount",
        "Customer Name", "Customer Phone", "Country Code"
    ];

    const headerRow = ws.addRow(header);
    headerRow.eachCell(cell => {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF000000" } };
        cell.alignment = { vertical: "middle", horizontal: "center" };
    });

    let totals = {
        slotCost: 0,
        finalAmount: 0,
        couponDiscount: 0,
        razorpay: 0,
        cash: 0,
        upi: 0,
        card: 0,
        wallet: 0,
        totalPaid: 0,
        balance: 0,
        commissionAmount: 0
    };

    bookings.forEach(b => {
        const payments = splitPayments(b.payment || []);

        const isCancelled = (b.status || "").toLowerCase() === "cancelled";
        // Online money refunded to the customer's GameOn Wallet is no longer
        // revenue — drop the Razorpay + Wallet portions so the report (and its
        // commission) matches the settlement/finance behaviour.
        if (isRefundedToWallet(b)) {
            payments.razorpay = 0;
            payments.wallet = 0;
        }
        const totalPaid = payments.razorpay + payments.cash + payments.upi + payments.card + payments.wallet;

        const couponDiscount = Number(b.coupon_discount) || 0;
        const finalAmount = isCancelled ? 0 : (Number(b.final_amount) || 0);
        const balance = isCancelled ? 0 : finalAmount - totalPaid;

        // Commission is charged on the online money received. For active bookings
        // that is the final amount; for cancelled ones it is whatever was kept
        // online (Razorpay + GameOn Wallet).
        const commissionAmount = isCancelled
            ? (((payments.razorpay + payments.wallet) * commission_percentage) / 100)
            : ((finalAmount * commission_percentage) / 100);

        totals.slotCost += Number(b.slot_cost) || 0;
        totals.finalAmount += finalAmount;
        totals.couponDiscount += couponDiscount;
        totals.razorpay += payments.razorpay;
        totals.cash += payments.cash;
        totals.upi += payments.upi;
        totals.card += payments.card;
        totals.wallet += payments.wallet;
        totals.totalPaid += totalPaid;
        totals.balance += balance;
        totals.commissionAmount += commissionAmount;

        const row = ws.addRow([
            b.id,
            b.booking_type,
            b.date,
            b.status,
            b.venue_name,
            b.court_id,
            (b.sport && b.sport.name) ? b.sport.name : "",
            b.start_time,
            b.end_time,
            calculateDuration(b.start_time, b.end_time),
            b.slot_cost,
            finalAmount,
            b.coupon_code,
            couponDiscount,
            payments.razorpay,
            payments.cash,
            payments.upi,
            payments.card,
            payments.wallet,
            totalPaid,
            balance,
            commission_percentage,
            commissionAmount,
            (b.user && b.user.name) ? b.user.name : "",
            (b.user && b.user.phone) ? b.user.phone : "",
            (b.user && b.user.country_code) ? b.user.country_code : ""
        ]);

        row.getCell(3).numFmt = "yyyy-mm-dd";
        [11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23].forEach(i => {
            row.getCell(i).numFmt = "#,##0.00";
        });

        if (isCancelled) {
            row.eachCell(cell => {
                cell.font = { color: { argb: "FF888888" } };
            });
        }
    });

    const totalRow = ws.addRow([
        "TOTAL", "", "", "", "", "", "", "", "",
        "",
        totals.slotCost,
        totals.finalAmount,
        "",
        totals.couponDiscount,
        totals.razorpay,
        totals.cash,
        totals.upi,
        totals.card,
        totals.wallet,
        totals.totalPaid,
        totals.balance,
        "",
        totals.commissionAmount,
        "", "", ""
    ]);

    totalRow.eachCell(cell => {
        cell.font = { bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF3CD" } };
        cell.border = { top: { style: "thin" }, bottom: { style: "double" } };
    });

    [11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 23].forEach(i => {
        totalRow.getCell(i).numFmt = "#,##0.00";
    });

    ws.columns.forEach(col => {
        let maxLength = 0;
        col.eachCell({ includeEmpty: true }, cell => {
            maxLength = Math.max(maxLength, (cell.value || "").toString().length);
        });
        col.width = maxLength < 10 ? 10 : maxLength + 2;
    });
}
function createPaymentSheet(wb, sheetName, bookings) {
    const ps = wb.addWorksheet(sheetName);

    const payHeader = ps.addRow([
        "Booking ID", "Booking Type", "Booking Date", "Booking Status", "Payment Method",
        "Amount", "Collected By", "Datetime"
    ]);

    payHeader.eachCell(cell => {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF000000" } };
        cell.alignment = { horizontal: "center" };
    });

    let paymentTotals = {
        amount: 0,
    };

    bookings.forEach(b => {
        (b.payment || []).forEach(p => {

            const isCancelled = (b.status || "").toLowerCase() === "cancelled";

            paymentTotals.amount += Number(p.amount) || 0;

            const row = ps.addRow([
                b.id,
                b.booking_type,
                b.date,
                b.status,
                p.method || "",
                Number(p.amount) || 0,
                p.paid_user_name || "",
                p.datetime || ""
            ]);

            row.getCell(6).numFmt = "#,##0.00";

            if (isCancelled) {
                row.eachCell(cell => {
                    cell.font = { color: { argb: "FF888888" } };
                });
            }
        });
    });

    const paymentTotalRow = ps.addRow([
        "TOTAL", "", "", "", "",
        paymentTotals.amount,
        "", ""
    ]);

    paymentTotalRow.eachCell(cell => {
        cell.font = { bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF3CD" } };
        cell.border = { top: { style: "thin" }, bottom: { style: "double" } };
    });

    [6].forEach(i => {
        paymentTotalRow.getCell(i).numFmt = "#,##0.00";
    });

    ps.columns.forEach(col => {
        let maxLength = 0;
        col.eachCell({ includeEmpty: true }, cell => {
            maxLength = Math.max(maxLength, (cell.value || "").toString().length);
        });
        col.width = maxLength < 12 ? 12 : maxLength + 2;
    });
}

async function sendEmail(venueName, toEmail) {
    let start, end;

    if (StatementType === "Monthly") {
        ({ start, end } = getPreviousMonthRange());
    } else {
        ({ start, end } = getPreviousWeekRange());
    }

    let transporter = nodemailer.createTransport({
        service: "gmail",
        auth: {
            user: process.env.EMAIL_USER,
            pass: process.env.EMAIL_PASS
        }
    });

    await transporter.sendMail({
        from: process.env.EMAIL_USER,
        to: toEmail,
        bcc: "rahmanvapu@gmail.com, abdras157@gmail.com",
        subject: `${venueName} - ${StatementType} Statement - ${start} to ${end}`,
        text: `Hello,

        Please find attached the ${StatementType.toLowerCase()} booking and payment report for your venue.

        This report includes:
        • Booking details
        • Online bookings
        • Payment breakdown

        If you have any questions or need clarification on any of the data, feel free to reach out.

        Thank you for your continued support.

        Best regards,
        GAMEON Team
        `,
        attachments: [{
            filename: `${StatementType}Statement.xlsx`,
            path: `./${StatementType}Statement.xlsx`
        }]
    });

    console.log("Email sent successfully!");
}


async function fetchVenueDetails() {

    const snapshot = await db
        .collection("venue_details")
        .where("status", "==", true)
        .orderBy("sequence", "asc")
        .get();

    return snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
    }));
}
async function fetchBookings(venueId) {
    let start, end;

    if (StatementType === "Monthly") {
        ({ start, end } = getPreviousMonthRange());
    } else {
        ({ start, end } = getPreviousWeekRange());
    }

    const snapshot = await db
        .collection("bookings")
        .where("venue_id", "==", venueId)
        .where("date", ">=", start)
        .where("date", "<=", end)
        .where("booking_type", "in", ["blocked", "offline", "online"])
        .get();

    return snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
    }));
}


async function runJob() {
    const venueDetails = await fetchVenueDetails();

    for (const venue of venueDetails) {
        //if (venue.id === "EpPEDwMbrODVupcyhohe") {
        const bookings = await fetchBookings(venue.id);

        if (bookings.length > 0) {
            const wb = new ExcelJS.Workbook();

            createBookingSheet(wb, "Booking Details", bookings);

            const onlineBookings = bookings.filter(
                b => (b.booking_type || "").toUpperCase() === "ONLINE"
            );
            createOnlineBookingSheet(wb, "Online Booking", onlineBookings, venue.commission_percentage);

            createPaymentSheet(wb, "Payment Details", bookings);

            await wb.xlsx.writeFile(`${StatementType}Statement.xlsx`);

            //await sendEmail(venue.name, 'abdras157@gmail.com');
            await sendEmail(venue.name, venue.email_address);
        }
        //}
    }
}

// Only initialize Firebase and run the job when executed directly
// (`node monthly_report.js`). When required as a module, the pure sheet-building
// helpers below are exported for testing without any credentials.
if (require.main === module) {
    const serviceAccount = JSON.parse(
        Buffer.from(process.env.FIREBASE_SERVICE_ACCOUNT, 'base64').toString('utf-8')
    );
    admin.initializeApp({
        credential: admin.credential.cert(serviceAccount)
    });
    db = admin.firestore();
    runJob();
}

module.exports = {
    splitPayments,
    isRefundedToWallet,
    createBookingSheet,
    createOnlineBookingSheet,
    createPaymentSheet,
};