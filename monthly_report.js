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
function totalPayments(payments) {
    return payments.razorpay + payments.cash + payments.upi + payments.card + payments.wallet;
}
function getBookingProductAmount(booking) {
    const storedAmount = Number(booking.product_amount);
    if (Number.isFinite(storedAmount)) return storedAmount;
    return (booking.products || []).reduce(
        (sum, product) => sum + (Number(product.total) || 0),
        0
    );
}
function getRecognizedPayments(booking) {
    const payments = splitPayments(booking.payment || []);
    if (isRefundedToWallet(booking)) {
        payments.razorpay = 0;
        payments.wallet = 0;
    }
    return payments;
}
function isOnlinePaymentMethod(method) {
    const normalized = String(method || "").trim().toLowerCase();
    return normalized.includes("razor") || normalized === "online" || normalized.includes("wallet");
}
function styleHeader(row) {
    row.eachCell(cell => {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF000000" } };
        cell.alignment = { vertical: "middle", horizontal: "center" };
    });
}
function styleTotalRow(row) {
    row.eachCell(cell => {
        cell.font = { bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF3CD" } };
        cell.border = { top: { style: "thin" }, bottom: { style: "double" } };
    });
}
function autoFitColumns(ws, minimumWidth = 10) {
    ws.columns.forEach(col => {
        let maxLength = 0;
        col.eachCell({ includeEmpty: true }, cell => {
            maxLength = Math.max(maxLength, (cell.value || "").toString().length);
        });
        col.width = maxLength < minimumWidth ? minimumWidth : Math.min(maxLength + 2, 60);
    });
    ws.views = [{ state: "frozen", ySplit: 1, showGridLines: false }];
    if (ws.columnCount > 0) {
        ws.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: 1, column: ws.columnCount },
        };
    }
}
function addPaymentSplits(target, source) {
    target.razorpay += source.razorpay;
    target.cash += source.cash;
    target.upi += source.upi;
    target.card += source.card;
    target.wallet += source.wallet;
    return target;
}
function calculateAnalyticsSummary(bookings, productSales = []) {
    const realBookings = bookings.filter(booking => !booking.is_dummy);
    const validBookings = realBookings.filter(
        booking => (booking.status || "").toLowerCase() !== "cancelled"
    );
    const validProductSales = productSales.filter(
        sale => (sale.payment_status || "").toLowerCase() !== "cancelled"
    );

    const slotRevenue = validBookings
        .filter(booking => booking.type !== "membership")
        .reduce((sum, booking) => sum + (Number(booking.final_amount) || 0), 0);
    const bookingProductRevenue = validBookings
        .filter(booking => booking.type !== "membership")
        .reduce((sum, booking) => sum + getBookingProductAmount(booking), 0);
    const standaloneProductRevenue = validProductSales.reduce(
        (sum, sale) => sum + (Number(sale.total_amount) || 0),
        0
    );
    const membershipRevenue = validBookings
        .filter(booking => booking.type === "membership")
        .reduce((sum, booking) => sum + (Number(booking.final_amount) || 0), 0);
    const retainedCancelledRevenue = realBookings
        .filter(booking => (booking.status || "").toLowerCase() === "cancelled")
        .reduce((sum, booking) => sum + totalPayments(getRecognizedPayments(booking)), 0);
    const productRevenue = bookingProductRevenue + standaloneProductRevenue;
    const totalRevenue = slotRevenue + membershipRevenue + retainedCancelledRevenue;
    const totalBusinessRevenue = totalRevenue + productRevenue;

    const payments = { razorpay: 0, cash: 0, upi: 0, card: 0, wallet: 0, unclassified: 0 };
    realBookings.forEach(booking => {
        addPaymentSplits(payments, getRecognizedPayments(booking));
        if (
            booking.type === "membership" &&
            booking.is_legacy_membership_row &&
            (!Array.isArray(booking.payment) || booking.payment.length === 0)
        ) {
            payments.unclassified += Number(booking.paid_amount) || 0;
        }
    });
    validProductSales.forEach(sale => {
        addPaymentSplits(payments, splitPayments(sale.payment || []));
    });

    const bookingDue = validBookings.reduce((sum, booking) => {
        const charge = (Number(booking.final_amount) || 0) +
            (booking.type === "membership" ? 0 : getBookingProductAmount(booking));
        return sum + Math.max(charge - totalPayments(getRecognizedPayments(booking)), 0);
    }, 0);
    const standaloneProductDue = validProductSales.reduce((sum, sale) => {
        const paid = totalPayments(splitPayments(sale.payment || []));
        return sum + Math.max((Number(sale.total_amount) || 0) - paid, 0);
    }, 0);

    return {
        totalBusinessRevenue,
        totalRevenue,
        slotRevenue,
        productRevenue,
        bookingProductRevenue,
        standaloneProductRevenue,
        membershipRevenue,
        retainedCancelledRevenue,
        payments,
        totalRecognizedPayments: totalPayments(payments) + payments.unclassified,
        pendingAmount: bookingDue + standaloneProductDue,
        totalBookings: realBookings.length,
        cancelledBookings: realBookings.filter(
            booking => (booking.status || "").toLowerCase() === "cancelled"
        ).length,
        blockedBookings: realBookings.filter(
            booking => (booking.status || "").toLowerCase() === "blocked"
        ).length,
        couponDiscount: validBookings.reduce(
            (sum, booking) => sum + (Number(booking.coupon_discount) || 0),
            0
        ),
        coinDiscount: validBookings.reduce(
            (sum, booking) => sum + (Number(booking.coin_discount) || 0),
            0
        ),
    };
}
function createAnalyticsSummarySheet(wb, sheetName, bookings, productSales = []) {
    const ws = wb.addWorksheet(sheetName);
    const summary = calculateAnalyticsSummary(bookings, productSales);
    styleHeader(ws.addRow(["Metric", "Amount / Count", "Definition"]));

    const rows = [
        ["Total Business Revenue", summary.totalBusinessRevenue, "Slot + products + memberships + retained cancelled booking payments"],
        ["Total Revenue", summary.totalRevenue, "Booking and membership revenue, plus retained cancelled booking payments"],
        ["Slot Revenue", summary.slotRevenue, "Final slot amount from non-cancelled, non-membership bookings"],
        ["Product Revenue", summary.productRevenue, "Booking product sales + standalone product sales"],
        ["Booking Product Revenue", summary.bookingProductRevenue, "Product charges attached to non-cancelled bookings"],
        ["Standalone Product Revenue", summary.standaloneProductRevenue, "Non-cancelled standalone product sales"],
        ["Membership Revenue", summary.membershipRevenue, "Final membership amount allocated to service dates in this month"],
        ["Retained Cancelled Revenue", summary.retainedCancelledRevenue, "Cancelled booking payments retained as revenue; wallet refunds excluded"],
        ["Online Payment", summary.payments.razorpay, "Recognized Online/Razorpay payments, including product sales"],
        ["Cash Payment", summary.payments.cash, "Recognized cash payments, including product sales"],
        ["UPI Payment", summary.payments.upi, "Recognized UPI payments, including product sales"],
        ["Card Payment", summary.payments.card, "Recognized card payments, including product sales"],
        ["GameOn Wallet Payment", summary.payments.wallet, "Recognized GameOn Wallet payments"],
        ["Legacy / Unclassified Payment", summary.payments.unclassified, "Allocated legacy paid amount without a stored payment method"],
        ["Total Recognized Payments", summary.totalRecognizedPayments, "Analytics amount from Payment Details"],
        ["Pending Amount", summary.pendingAmount, "Remaining booking, membership, and standalone product-sale due"],
        ["Coupon Discount", summary.couponDiscount, "Coupon discounts on non-cancelled bookings"],
        ["Coin Discount", summary.coinDiscount, "Coin discounts on non-cancelled bookings"],
        ["Total Bookings", summary.totalBookings, "Dated bookings plus allocated legacy membership dates"],
        ["Cancelled Bookings", summary.cancelledBookings, "Cancelled booking rows"],
        ["Blocked Bookings", summary.blockedBookings, "Blocked booking rows"],
    ];

    rows.forEach((values, index) => {
        const row = ws.addRow(values);
        row.getCell(2).numFmt = index >= 18 ? "#,##0" : "#,##0.00";
    });
    ws.getColumn(1).font = { bold: true };
    autoFitColumns(ws, 16);
}
function createBookingSheet(wb, sheetName, bookings) {
    const ws = wb.addWorksheet(sheetName);

    const header = [
        "Booking ID", "Booking Type", "Booking Date", "Booking Status",
        "Venue Name", "Court ID", "Sport", "Start Time", "End Time",
        "Duration", "Slot Cost", "Final Amount", "Product Sales",
        "Coupon Code", "Coupon Discount",
        "Online", "Cash", "UPI", "Card", "GameOn Wallet",
        "Total Paid", "Balance/Due",
        "Customer Name", "Customer Phone", "Country Code"
    ];

    const headerRow = ws.addRow(header);
    styleHeader(headerRow);

    let totals = {
        slotCost: 0,
        finalAmount: 0,
        productAmount: 0,
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
        const isCancelled = (b.status || "").toLowerCase() === "cancelled";
        const payments = getRecognizedPayments(b);
        const totalPaid = totalPayments(payments);

        const couponDiscount = Number(b.coupon_discount) || 0;
        const finalAmount = isCancelled ? 0 : (Number(b.final_amount) || 0);
        const productAmount = isCancelled ? 0 : getBookingProductAmount(b);
        const balance = isCancelled ? 0 : Math.max(finalAmount + productAmount - totalPaid, 0);

        totals.slotCost += Number(b.slot_cost) || 0;
        totals.finalAmount += finalAmount;
        totals.productAmount += productAmount;
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
            productAmount,
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
        [11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22].forEach(i => {
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
        totals.productAmount,
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

    styleTotalRow(totalRow);

    [11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22].forEach(i => {
        totalRow.getCell(i).numFmt = "#,##0.00";
    });

    autoFitColumns(ws);
}
function createOnlineBookingSheet(wb, sheetName, bookings, commission_percentage) {
    const ws = wb.addWorksheet(sheetName);

    const header = [
        "Booking ID", "Booking Type", "Booking Date", "Booking Status",
        "Venue Name", "Court ID", "Sport", "Start Time", "End Time",
        "Duration", "Slot Cost", "Final Amount", "Product Sales",
        "Coupon Code", "Coupon Discount",
        "Online", "Cash", "UPI", "Card", "GameOn Wallet",
        "Total Paid", "Balance/Due", "Commission Percentage", "Commission Amount",
        "Customer Name", "Customer Phone", "Country Code"
    ];

    const headerRow = ws.addRow(header);
    styleHeader(headerRow);

    let totals = {
        slotCost: 0,
        finalAmount: 0,
        productAmount: 0,
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
        const isCancelled = (b.status || "").toLowerCase() === "cancelled";
        const payments = getRecognizedPayments(b);
        const totalPaid = totalPayments(payments);

        const couponDiscount = Number(b.coupon_discount) || 0;
        const finalAmount = isCancelled ? 0 : (Number(b.final_amount) || 0);
        const productAmount = isCancelled ? 0 : getBookingProductAmount(b);
        const balance = isCancelled ? 0 : Math.max(finalAmount + productAmount - totalPaid, 0);

        // Commission is charged on the online money received. For active bookings
        // that is the final amount; for cancelled ones it is whatever was kept
        // online (Razorpay + GameOn Wallet).
        const commissionAmount = isCancelled
            ? (((payments.razorpay + payments.wallet) * commission_percentage) / 100)
            : ((finalAmount * commission_percentage) / 100);

        totals.slotCost += Number(b.slot_cost) || 0;
        totals.finalAmount += finalAmount;
        totals.productAmount += productAmount;
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
            productAmount,
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
        [11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24].forEach(i => {
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
        totals.productAmount,
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

    styleTotalRow(totalRow);

    [11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 24].forEach(i => {
        totalRow.getCell(i).numFmt = "#,##0.00";
    });

    autoFitColumns(ws);
}
function createPaymentSheet(wb, sheetName, bookings, standaloneSales = []) {
    const ps = wb.addWorksheet(sheetName);

    const payHeader = ps.addRow([
        "Payment Source", "Reference ID", "Date", "Status", "Payment Method",
        "Recorded Amount", "Analytics Amount", "Collected By", "Datetime", "Notes"
    ]);

    styleHeader(payHeader);

    let paymentTotals = {
        recordedAmount: 0,
        analyticsAmount: 0,
    };

    bookings.forEach(b => {
        (b.payment || []).forEach(p => {
            const isCancelled = (b.status || "").toLowerCase() === "cancelled";
            const amount = Number(p.amount) || 0;
            const returnedToWallet = isRefundedToWallet(b) && isOnlinePaymentMethod(p.method);
            const analyticsAmount = returnedToWallet ? 0 : amount;
            const hasProducts = b.type !== "membership" && getBookingProductAmount(b) > 0;
            const source = b.type === "membership"
                ? "Membership"
                : hasProducts
                    ? "Booking + Products"
                    : "Booking";
            const notes = returnedToWallet
                ? "Returned to GameOn Wallet; excluded from analytics revenue"
                : hasProducts
                    ? "Payment covers booking slot and product charges"
                    : b.is_legacy_membership_row
                        ? "Legacy membership payment allocated to this date"
                        : "";

            paymentTotals.recordedAmount += amount;
            paymentTotals.analyticsAmount += analyticsAmount;

            const row = ps.addRow([
                source,
                b.report_membership_id || b.id,
                b.date,
                b.status,
                p.method || "",
                amount,
                analyticsAmount,
                p.paid_user_name || "",
                p.datetime || p.payment_datetime || "",
                notes,
            ]);

            [6, 7].forEach(index => {
                row.getCell(index).numFmt = "#,##0.00";
            });

            if (isCancelled) {
                row.eachCell(cell => {
                    cell.font = { color: { argb: "FF888888" } };
                });
            }
        });

        if (
            b.type === "membership" &&
            b.is_legacy_membership_row &&
            (!Array.isArray(b.payment) || b.payment.length === 0) &&
            Number(b.paid_amount) > 0
        ) {
            const amount = Number(b.paid_amount) || 0;
            const isCancelled = (b.status || "").toLowerCase() === "cancelled";
            paymentTotals.recordedAmount += amount;
            paymentTotals.analyticsAmount += amount;
            const row = ps.addRow([
                "Membership", b.report_membership_id || b.id, b.date, b.status,
                "Legacy / unspecified", amount, amount, "", "",
                "Legacy membership paid amount allocated to this date",
            ]);
            [6, 7].forEach(index => {
                row.getCell(index).numFmt = "#,##0.00";
            });
            if (isCancelled) {
                row.eachCell(cell => {
                    cell.font = { color: { argb: "FF888888" } };
                });
            }
        }
    });

    standaloneSales.forEach(sale => {
        (sale.payment || []).forEach(payment => {
            const amount = Number(payment.amount) || 0;
            const isCancelled = (sale.payment_status || "").toLowerCase() === "cancelled";
            const analyticsAmount = isCancelled ? 0 : amount;
            paymentTotals.recordedAmount += amount;
            paymentTotals.analyticsAmount += analyticsAmount;
            const row = ps.addRow([
                "Standalone Product Sale", sale.id, sale.date, sale.payment_status,
                payment.method || "", amount, analyticsAmount,
                payment.paid_user_name || sale.created_by?.name || "",
                payment.datetime || payment.payment_datetime || "",
                isCancelled ? "Cancelled product sale; excluded from analytics payments" : "",
            ]);
            [6, 7].forEach(index => {
                row.getCell(index).numFmt = "#,##0.00";
            });
            if (isCancelled) {
                row.eachCell(cell => {
                    cell.font = { color: { argb: "FF888888" } };
                });
            }
        });
    });

    const paymentTotalRow = ps.addRow([
        "TOTAL", "", "", "", "",
        paymentTotals.recordedAmount,
        paymentTotals.analyticsAmount,
        "", "", ""
    ]);

    paymentTotalRow.eachCell(cell => {
        cell.font = { bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF3CD" } };
        cell.border = { top: { style: "thin" }, bottom: { style: "double" } };
    });

    [6, 7].forEach(i => {
        paymentTotalRow.getCell(i).numFmt = "#,##0.00";
    });

    autoFitColumns(ps, 12);
}

function allocateBookingProductPayments(booking) {
    if ((booking.status || "").toLowerCase() === "cancelled") return [];

    let slotAmountRemaining = Number(booking.final_amount) || 0;
    let productAmountRemaining = getBookingProductAmount(booking);
    const allocated = [];

    for (const payment of booking.payment || []) {
        let amount = Number(payment.amount) || 0;
        if (amount <= 0) continue;

        const slotPortion = Math.min(amount, slotAmountRemaining);
        slotAmountRemaining = Math.max(slotAmountRemaining - slotPortion, 0);
        amount -= slotPortion;

        const productPortion = Math.min(amount, productAmountRemaining);
        if (productPortion > 0) {
            allocated.push({ ...payment, amount: productPortion });
            productAmountRemaining = Math.max(productAmountRemaining - productPortion, 0);
        }
        if (productAmountRemaining <= 0) break;
    }

    return allocated;
}

function formatProductItems(items = []) {
    return items
        .map(item => {
            const name = item.product_name || item.name || "Product";
            const quantity = Number(item.quantity) || 0;
            const price = Number(item.selling_price ?? item.price ?? item.cost) || 0;
            const total = Number(item.total) || price * quantity;
            return `${name} x ${quantity} @ ${price.toFixed(2)} = ${total.toFixed(2)}`;
        })
        .join("; ");
}

function createProductSalesSheet(wb, sheetName, bookings, standaloneSales = []) {
    const ws = wb.addWorksheet(sheetName);
    const header = [
        "Source", "Reference ID", "Date", "Time/Slot", "Status",
        "Venue Name", "Court", "Products", "Total Quantity", "Product Sales",
        "Online", "Cash", "UPI", "Card", "GameOn Wallet",
        "Paid Toward Products", "Balance/Due", "Customer Name", "Customer Phone",
        "Created/Collected By"
    ];
    styleHeader(ws.addRow(header));

    const rows = [];
    bookings
        .filter(
            booking =>
                !booking.is_dummy &&
                booking.type !== "membership" &&
                getBookingProductAmount(booking) > 0
        )
        .forEach(booking => {
            const isCancelled = (booking.status || "").toLowerCase() === "cancelled";
            const productAmount = isCancelled ? 0 : getBookingProductAmount(booking);
            const productPayments = splitPayments(allocateBookingProductPayments(booking));
            const paid = totalPayments(productPayments);
            rows.push({
                source: "Booking Add-on",
                referenceId: booking.id,
                date: booking.date,
                time: [booking.start_time, booking.end_time].filter(Boolean).join(" - "),
                status: booking.status,
                venueName: booking.venue_name,
                court: booking.court_name || booking.court_id,
                items: booking.products || [],
                productAmount,
                payments: productPayments,
                paid,
                due: isCancelled ? 0 : Math.max(productAmount - paid, 0),
                customerName: booking.user?.name || "",
                customerPhone: booking.user?.phone || "",
                collectedBy: (booking.payment || [])
                    .map(payment => payment.paid_user_name)
                    .filter(Boolean)
                    .filter((name, index, values) => values.indexOf(name) === index)
                    .join(", "),
                isCancelled,
            });
        });

    standaloneSales.forEach(sale => {
        const isCancelled = (sale.payment_status || "").toLowerCase() === "cancelled";
        const productAmount = isCancelled ? 0 : (Number(sale.total_amount) || 0);
        const payments = isCancelled
            ? { razorpay: 0, cash: 0, upi: 0, card: 0, wallet: 0 }
            : splitPayments(sale.payment || []);
        const paid = totalPayments(payments);
        rows.push({
            source: "Standalone Sale",
            referenceId: sale.id,
            date: sale.date,
            time: sale.sale_time || "",
            status: sale.payment_status,
            venueName: sale.venue_name || "",
            court: sale.court_name || sale.court_id || "",
            items: sale.items || [],
            productAmount,
            payments,
            paid,
            due: isCancelled ? 0 : Math.max(productAmount - paid, 0),
            customerName: sale.user?.name || sale.customer?.name || "",
            customerPhone: sale.user?.phone || sale.customer?.phone || "",
            collectedBy: sale.created_by?.name || (sale.payment || [])
                .map(payment => payment.paid_user_name)
                .filter(Boolean)
                .join(", "),
            isCancelled,
        });
    });

    rows.sort((a, b) => `${a.date || ""}-${a.time || ""}`.localeCompare(`${b.date || ""}-${b.time || ""}`));

    const totals = {
        quantity: 0, productAmount: 0, razorpay: 0, cash: 0, upi: 0,
        card: 0, wallet: 0, paid: 0, due: 0,
    };

    rows.forEach(item => {
        const quantity = item.items.reduce((sum, product) => sum + (Number(product.quantity) || 0), 0);
        totals.quantity += quantity;
        totals.productAmount += item.productAmount;
        totals.razorpay += item.payments.razorpay;
        totals.cash += item.payments.cash;
        totals.upi += item.payments.upi;
        totals.card += item.payments.card;
        totals.wallet += item.payments.wallet;
        totals.paid += item.paid;
        totals.due += item.due;

        const row = ws.addRow([
            item.source, item.referenceId, item.date, item.time, item.status,
            item.venueName, item.court, formatProductItems(item.items), quantity,
            item.productAmount, item.payments.razorpay, item.payments.cash,
            item.payments.upi, item.payments.card, item.payments.wallet,
            item.paid, item.due, item.customerName, item.customerPhone,
            item.collectedBy,
        ]);
        [10, 11, 12, 13, 14, 15, 16, 17].forEach(index => {
            row.getCell(index).numFmt = "#,##0.00";
        });
        if (item.isCancelled) {
            row.eachCell(cell => {
                cell.font = { color: { argb: "FF888888" } };
            });
        }
    });

    const totalRow = ws.addRow([
        "TOTAL", "", "", "", "", "", "", "", totals.quantity,
        totals.productAmount, totals.razorpay, totals.cash, totals.upi,
        totals.card, totals.wallet, totals.paid, totals.due, "", "", "",
    ]);
    styleTotalRow(totalRow);
    [10, 11, 12, 13, 14, 15, 16, 17].forEach(index => {
        totalRow.getCell(index).numFmt = "#,##0.00";
    });
    autoFitColumns(ws, 12);
}

function getDatesBetweenInclusive(start, end) {
    if (!start || !end) return [];
    const dates = [];
    const current = new Date(`${start}T00:00:00Z`);
    const last = new Date(`${end}T00:00:00Z`);
    if (Number.isNaN(current.getTime()) || Number.isNaN(last.getTime())) return dates;
    while (current <= last) {
        dates.push(current.toISOString().slice(0, 10));
        current.setUTCDate(current.getUTCDate() + 1);
    }
    return dates;
}

function splitAmountEqually(total, count) {
    if (!count || count <= 0) return [];
    const amount = Number(total) || 0;
    const base = Math.floor((amount / count) * 100) / 100;
    const result = new Array(count).fill(base);
    const remainder = Math.round((amount - base * count) * 100) / 100;
    result[result.length - 1] = Math.round((result[result.length - 1] + remainder) * 100) / 100;
    return result;
}

function allocatePaymentsToDates(payments, dates, perDateAmounts) {
    const allocations = Object.fromEntries(dates.map(date => [date, []]));
    const remaining = [...perDateAmounts];
    let dateIndex = 0;

    for (const payment of payments || []) {
        let amount = Number(payment.amount) || 0;
        while (amount > 0.001 && dateIndex < dates.length) {
            const allocated = Math.min(amount, remaining[dateIndex] || 0);
            if (allocated > 0) {
                allocations[dates[dateIndex]].push({
                    ...payment,
                    amount: Math.round(allocated * 100) / 100,
                });
                remaining[dateIndex] = Math.round((remaining[dateIndex] - allocated) * 100) / 100;
                amount = Math.round((amount - allocated) * 100) / 100;
            }
            if (remaining[dateIndex] <= 0.001 || allocated <= 0) dateIndex += 1;
        }
    }
    return allocations;
}

function allocateScalarToDates(total, dates, perDateAmounts) {
    let remaining = Number(total) || 0;
    return dates.map((date, index) => {
        const amount = Math.min(remaining, perDateAmounts[index] || 0);
        remaining = Math.round((remaining - amount) * 100) / 100;
        return Math.round(amount * 100) / 100;
    });
}

function buildMembershipReportRows(bookings, memberships, start, end) {
    const membershipById = new Map(
        memberships.filter(item => !item.is_dummy).map(item => [item.id, item])
    );
    const childRows = bookings
        .filter(booking => booking.type === "membership" && !booking.is_dummy)
        .map(booking => {
            const membershipId = booking.membership_id || booking.membership_parent_id || "";
            const parent = membershipById.get(membershipId) || {};
            return {
                ...booking,
                report_membership_id: membershipId || booking.id,
                membership_start_date: parent.start_date || booking.membership_start_date || "",
                membership_end_date: parent.end_date || booking.membership_end_date || "",
            };
        });

    const legacyRows = [];
    memberships
        .filter(membership => !membership.is_dummy && !membership.has_child_bookings)
        .filter(membership => membership.start_date <= end && membership.end_date >= start)
        .forEach(membership => {
            const excluded = new Set(membership.excluded_dates || []);
            const activeDates = getDatesBetweenInclusive(membership.start_date, membership.end_date)
                .filter(date => !excluded.has(date));
            const finalAmounts = splitAmountEqually(
                membership.final_amount ?? membership.slot_cost,
                activeDates.length
            );
            const slotAmounts = splitAmountEqually(
                membership.slot_cost ?? membership.final_amount,
                activeDates.length
            );
            const paymentsByDate = allocatePaymentsToDates(
                membership.payment || [],
                activeDates,
                finalAmounts
            );
            const paidByDate = allocateScalarToDates(
                membership.paid_amount,
                activeDates,
                finalAmounts
            );

            activeDates.forEach((date, index) => {
                if (date < start || date > end) return;
                const payments = paymentsByDate[date] || [];
                const paidFromPayments = payments.reduce(
                    (sum, payment) => sum + (Number(payment.amount) || 0),
                    0
                );
                legacyRows.push({
                    ...membership,
                    date,
                    slot_cost: slotAmounts[index] || 0,
                    final_amount: finalAmounts[index] || 0,
                    payment: payments,
                    paid_amount: payments.length ? paidFromPayments : paidByDate[index] || 0,
                    report_membership_id: membership.id,
                    membership_start_date: membership.start_date,
                    membership_end_date: membership.end_date,
                    is_legacy_membership_row: true,
                });
            });
        });

    return [...childRows, ...legacyRows].sort((a, b) =>
        `${a.date || ""}-${a.start_time || ""}`.localeCompare(`${b.date || ""}-${b.start_time || ""}`)
    );
}

function createMembershipSheet(wb, sheetName, membershipRows) {
    const ws = wb.addWorksheet(sheetName);
    const header = [
        "Membership ID", "Booking Date", "Membership Start", "Membership End",
        "Status", "Venue Name", "Court", "Sport", "Start Time", "End Time",
        "Allocated Original Slot Value", "Allocated Membership Revenue", "Allocated Discount",
        "Online", "Cash", "UPI", "Card", "GameOn Wallet", "Allocated Paid", "Balance/Due",
        "Customer Name", "Customer Phone", "Record Type"
    ];
    styleHeader(ws.addRow(header));

    const totals = {
        slot: 0, revenue: 0, discount: 0, razorpay: 0, cash: 0, upi: 0, card: 0,
        wallet: 0, paid: 0, due: 0,
    };

    membershipRows.forEach(membership => {
        const isCancelled = (membership.status || "").toLowerCase() === "cancelled";
        const slotAmount = isCancelled ? 0 : (Number(membership.slot_cost) || 0);
        const revenue = isCancelled ? 0 : (Number(membership.final_amount) || 0);
        const discount = Math.max(slotAmount - revenue, 0);
        const payments = splitPayments(membership.payment || []);
        const paid = payments.razorpay || payments.cash || payments.upi || payments.card || payments.wallet
            ? totalPayments(payments)
            : (Number(membership.paid_amount) || 0);
        const due = isCancelled ? 0 : Math.max(revenue - paid, 0);

        totals.slot += slotAmount;
        totals.revenue += revenue;
        totals.discount += discount;
        totals.razorpay += payments.razorpay;
        totals.cash += payments.cash;
        totals.upi += payments.upi;
        totals.card += payments.card;
        totals.wallet += payments.wallet;
        totals.paid += paid;
        totals.due += due;

        const row = ws.addRow([
            membership.report_membership_id, membership.date,
            membership.membership_start_date, membership.membership_end_date,
            membership.status, membership.venue_name,
            membership.court_name || membership.court_id,
            membership.sport?.name || "", membership.start_time, membership.end_time,
            slotAmount, revenue, discount, payments.razorpay, payments.cash, payments.upi,
            payments.card, payments.wallet, paid, due,
            membership.user?.name || "", membership.user?.phone || "",
            membership.is_legacy_membership_row ? "Legacy allocated" : "Per-date booking",
        ]);
        [11, 12, 13, 14, 15, 16, 17, 18, 19, 20].forEach(index => {
            row.getCell(index).numFmt = "#,##0.00";
        });
        if (isCancelled) {
            row.eachCell(cell => {
                cell.font = { color: { argb: "FF888888" } };
            });
        }
    });

    const totalRow = ws.addRow([
        "TOTAL", "", "", "", "", "", "", "", "", "",
        totals.slot, totals.revenue, totals.discount, totals.razorpay, totals.cash, totals.upi,
        totals.card, totals.wallet, totals.paid, totals.due, "", "", "",
    ]);
    styleTotalRow(totalRow);
    [11, 12, 13, 14, 15, 16, 17, 18, 19, 20].forEach(index => {
        totalRow.getCell(index).numFmt = "#,##0.00";
    });
    autoFitColumns(ws, 12);
}

async function sendEmail(venueName, toEmail, start, end) {
    let transporter = nodemailer.createTransport({
        service: "gmail",
        auth: {
            user: process.env.EMAIL_USER,
            pass: process.env.EMAIL_PASS
        }
    });

    const message = {
        from: process.env.EMAIL_USER,
        to: toEmail,
        subject: `${venueName} - ${StatementType} Statement - ${start} to ${end}`,
        text: `Hello,

        Please find attached the ${StatementType.toLowerCase()} booking and payment report for your venue.

        This report includes:
        • Analytics summary
        • Booking details
        • Online bookings
        • Product sales
        • Membership allocations for this report period
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
    };

    // A manual test override must send only to its requested recipient.
    // Normal scheduled deliveries retain the existing operational BCC list.
    if (!process.env.REPORT_EMAIL_OVERRIDE) {
        message.bcc = "rahmanvapu@gmail.com, abdras157@gmail.com";
    }

    await transporter.sendMail(message);

    console.log(`Email sent successfully for ${venueName}.`);
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
async function fetchBookings(venueId, start, end) {
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

async function fetchProductSales(venueId, start, end) {
    const snapshot = await db
        .collection("product_sales")
        .where("venue_id", "==", venueId)
        .where("date", ">=", start)
        .where("date", "<=", end)
        .get();

    return snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
}

async function fetchMemberships(venueId) {
    const snapshot = await db
        .collection("memberships")
        .where("venue_id", "==", venueId)
        .get();

    return snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
}


async function runJob() {
    const range = StatementType === "Monthly"
        ? getPreviousMonthRange()
        : getPreviousWeekRange();
    const venueFilter = String(process.env.REPORT_VENUE_NAME || "").trim();
    const emailOverride = String(process.env.REPORT_EMAIL_OVERRIDE || "").trim();

    if (Boolean(venueFilter) !== Boolean(emailOverride)) {
        throw new Error(
            "REPORT_VENUE_NAME and REPORT_EMAIL_OVERRIDE must be provided together."
        );
    }

    let venueDetails = await fetchVenueDetails();
    if (venueFilter) {
        venueDetails = venueDetails.filter(
            venue => String(venue.name || "").trim().toLowerCase() === venueFilter.toLowerCase()
        );
        if (venueDetails.length !== 1) {
            throw new Error(`Expected exactly one active venue named \"${venueFilter}\".`);
        }
    }

    for (const venue of venueDetails) {
        const [bookings, productSales, memberships] = await Promise.all([
            fetchBookings(venue.id, range.start, range.end),
            fetchProductSales(venue.id, range.start, range.end),
            fetchMemberships(venue.id),
        ]);
        const reportBookings = bookings.filter(booking => !booking.is_dummy);
        const membershipRows = buildMembershipReportRows(
            reportBookings,
            memberships,
            range.start,
            range.end
        );
        const legacyMembershipRows = membershipRows.filter(
            membership => membership.is_legacy_membership_row
        );
        // This mirrors the Analytics screen, which combines dated bookings with
        // date-allocated legacy memberships. New membership child bookings are
        // already present in reportBookings and are not added twice.
        const analyticsBookings = [...reportBookings, ...legacyMembershipRows];

        if (reportBookings.length > 0 || productSales.length > 0 || membershipRows.length > 0) {
            const wb = new ExcelJS.Workbook();

            createAnalyticsSummarySheet(
                wb,
                "Analytics Summary",
                analyticsBookings,
                productSales
            );
            createBookingSheet(wb, "Booking Details", analyticsBookings);

            const onlineBookings = analyticsBookings.filter(
                b => (b.booking_type || "").toUpperCase() === "ONLINE"
            );
            createOnlineBookingSheet(wb, "Online Booking", onlineBookings, venue.commission_percentage);

            createProductSalesSheet(
                wb,
                "Product Sales",
                analyticsBookings,
                productSales.map(sale => ({
                    ...sale,
                    venue_name: sale.venue_name || venue.name,
                }))
            );
            createMembershipSheet(wb, "Memberships", membershipRows);
            createPaymentSheet(
                wb,
                "Payment Details",
                analyticsBookings,
                productSales
            );

            await wb.xlsx.writeFile(`${StatementType}Statement.xlsx`);

            await sendEmail(
                venue.name,
                emailOverride || venue.email_address,
                range.start,
                range.end
            );
        } else {
            console.log(`No reportable activity for ${venue.name}.`);
        }
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
    runJob().catch(error => {
        console.error("Monthly report failed:", error);
        process.exitCode = 1;
    });
}

module.exports = {
    allocateBookingProductPayments,
    buildMembershipReportRows,
    calculateAnalyticsSummary,
    createAnalyticsSummarySheet,
    createBookingSheet,
    createMembershipSheet,
    createOnlineBookingSheet,
    createPaymentSheet,
    createProductSalesSheet,
    getBookingProductAmount,
    splitPayments,
    isRefundedToWallet,
};
