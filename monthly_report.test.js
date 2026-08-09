const test = require("node:test");
const assert = require("node:assert/strict");
const ExcelJS = require("exceljs");

const {
    buildMembershipReportRows,
    calculateAnalyticsSummary,
    createBookingSheet,
    createMembershipSheet,
    createOnlineBookingSheet,
    createPaymentSheet,
    createProductSalesSheet,
    isRefundedToWallet,
    splitPayments,
} = require("./monthly_report");

const onlineBooking = (overrides = {}) => ({
    id: "booking-1",
    booking_type: "online",
    status: "confirmed",
    final_amount: 1000,
    payment: [{ method: "Online", amount: 1000 }],
    ...overrides,
});

const getOnlineReportRow = (booking, commissionPercentage = 10) => {
    const workbook = new ExcelJS.Workbook();
    createOnlineBookingSheet(
        workbook,
        "Online Booking",
        [booking],
        commissionPercentage,
    );
    return workbook.getWorksheet("Online Booking").getRow(2);
};

test("recognizes current online and GameOn Wallet payment labels", () => {
    assert.deepEqual(
        splitPayments([
            { method: "Online", amount: 100 },
            { method: "Razorpay", amount: 200 },
            { method: "GameOn Wallet", amount: 300 },
            { method: "Cash", amount: 400 },
        ]),
        { razorpay: 300, cash: 400, upi: 0, card: 0, wallet: 300 },
    );
});

test("keeps the legacy cancellation calculation when no wallet-refund fields exist", () => {
    const booking = onlineBooking({ status: "cancelled" });
    const row = getOnlineReportRow(booking);

    assert.equal(isRefundedToWallet(booking), false);
    assert.equal(row.getCell(12).value, 0, "cancelled slot revenue");
    assert.equal(row.getCell(16).value, 1000, "legacy online payment");
    assert.equal(row.getCell(21).value, 1000, "legacy total paid");
    assert.equal(row.getCell(24).value, 100, "legacy commission");
});

test("excludes a new wallet-refunded cancellation from revenue and commission", () => {
    const booking = onlineBooking({
        status: "cancelled",
        wallet_refund_choice: "refund_to_gameon_wallet",
        wallet_refund_amount: 1000,
    });
    const row = getOnlineReportRow(booking);

    assert.equal(isRefundedToWallet(booking), true);
    assert.equal(row.getCell(12).value, 0, "cancelled slot revenue");
    assert.equal(row.getCell(16).value, 0, "recognized online payment");
    assert.equal(row.getCell(20).value, 0, "recognized wallet payment");
    assert.equal(row.getCell(21).value, 0, "recognized total paid");
    assert.equal(row.getCell(24).value, 0, "commission");
});

test("includes booking product sales when calculating balance due", () => {
    const booking = onlineBooking({
        final_amount: 1000,
        product_amount: 250,
        products: [{ product_name: "Water", quantity: 5, price: 50, total: 250 }],
        payment: [
            { method: "Online", amount: 1000 },
            { method: "Cash", amount: 250, paid_user_name: "Manager" },
        ],
    });
    const workbook = new ExcelJS.Workbook();
    createBookingSheet(workbook, "Booking Details", [booking]);
    createOnlineBookingSheet(workbook, "Online Booking", [booking], 10);

    const bookingRow = workbook.getWorksheet("Booking Details").getRow(2);
    const onlineRow = workbook.getWorksheet("Online Booking").getRow(2);
    assert.equal(bookingRow.getCell(13).value, 250, "product sales");
    assert.equal(bookingRow.getCell(21).value, 1250, "total paid");
    assert.equal(bookingRow.getCell(22).value, 0, "balance due");
    assert.equal(onlineRow.getCell(24).value, 100, "slot-only commission");
});

test("creates product sales rows for booking add-ons and standalone sales", () => {
    const booking = onlineBooking({
        final_amount: 1000,
        product_amount: 250,
        products: [{ product_name: "Water", quantity: 5, price: 50, total: 250 }],
        payment: [
            { method: "Online", amount: 1000 },
            { method: "Cash", amount: 250, paid_user_name: "Manager" },
        ],
    });
    const standalone = {
        id: "sale-1",
        date: "2026-07-15",
        payment_status: "partial",
        total_amount: 200,
        paid_amount: 100,
        items: [{ product_name: "Grip", quantity: 2, selling_price: 100, total: 200 }],
        payment: [{ method: "UPI", amount: 100 }],
        created_by: { name: "Manager" },
    };
    const workbook = new ExcelJS.Workbook();
    createProductSalesSheet(workbook, "Product Sales", [booking], [standalone]);
    const sheet = workbook.getWorksheet("Product Sales");

    assert.equal(sheet.getRow(2).getCell(1).value, "Booking Add-on");
    assert.equal(sheet.getRow(2).getCell(12).value, 250, "cash allocated to products");
    assert.equal(sheet.getRow(2).getCell(17).value, 0, "booking product due");
    assert.equal(sheet.getRow(3).getCell(1).value, "Standalone Sale");
    assert.equal(sheet.getRow(3).getCell(13).value, 100, "standalone UPI paid");
    assert.equal(sheet.getRow(3).getCell(17).value, 100, "standalone due");
});

test("allocates a multi-month legacy membership only to dates in the report month", () => {
    const membership = {
        id: "membership-1",
        type: "membership",
        status: "active",
        start_date: "2026-07-31",
        end_date: "2026-08-02",
        slot_cost: 360,
        final_amount: 300,
        paid_amount: 300,
        payment: [{ method: "Cash", amount: 300 }],
    };

    const rows = buildMembershipReportRows(
        [],
        [membership],
        "2026-08-01",
        "2026-08-31",
    );

    assert.deepEqual(rows.map(row => row.date), ["2026-08-01", "2026-08-02"]);
    assert.deepEqual(rows.map(row => row.final_amount), [100, 100]);
    assert.deepEqual(rows.map(row => row.slot_cost), [120, 120]);
    assert.deepEqual(rows.map(row => row.paid_amount), [100, 100]);
});

test("labels membership original value, revenue, and discount clearly", () => {
    const workbook = new ExcelJS.Workbook();
    createMembershipSheet(workbook, "Memberships", [{
        report_membership_id: "membership-1",
        date: "2026-08-01",
        membership_start_date: "2026-07-01",
        membership_end_date: "2026-08-31",
        status: "active",
        slot_cost: 150,
        final_amount: 120,
        paid_amount: 120,
        payment: [{ method: "Cash", amount: 120 }],
    }]);
    const sheet = workbook.getWorksheet("Memberships");

    assert.equal(sheet.getRow(1).getCell(11).value, "Allocated Original Slot Value");
    assert.equal(sheet.getRow(1).getCell(12).value, "Allocated Membership Revenue");
    assert.equal(sheet.getRow(1).getCell(13).value, "Allocated Discount");
    assert.equal(sheet.getRow(2).getCell(13).value, 30);
});

test("payment details includes booking products, memberships, product sales, and recognition adjustments", () => {
    const productBooking = onlineBooking({
        product_amount: 250,
        payment: [
            { method: "Online", amount: 1000 },
            { method: "Cash", amount: 250 },
        ],
    });
    const membership = {
        id: "membership-row",
        report_membership_id: "membership-1",
        type: "membership",
        status: "active",
        date: "2026-07-10",
        payment: [{ method: "Cash", amount: 100 }],
        is_legacy_membership_row: true,
    };
    const walletRefund = onlineBooking({
        id: "wallet-refund",
        status: "cancelled",
        wallet_refund_choice: "refund_to_gameon_wallet",
        wallet_refund_amount: 100,
        payment: [{ method: "Online", amount: 100 }],
    });
    const productSales = [
        {
            id: "sale-active",
            date: "2026-07-11",
            payment_status: "partial",
            payment: [{ method: "UPI", amount: 100 }],
        },
        {
            id: "sale-cancelled",
            date: "2026-07-12",
            payment_status: "cancelled",
            payment: [{ method: "Cash", amount: 80 }],
        },
    ];
    const workbook = new ExcelJS.Workbook();
    createPaymentSheet(
        workbook,
        "Payment Details",
        [productBooking, membership, walletRefund],
        productSales,
    );
    const sheet = workbook.getWorksheet("Payment Details");
    const dataRows = [];
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1 && row.getCell(1).value !== "TOTAL") dataRows.push(row);
    });

    assert.deepEqual(
        [...new Set(dataRows.map(row => row.getCell(1).value))],
        ["Booking + Products", "Membership", "Booking", "Standalone Product Sale"],
    );
    const walletRow = dataRows.find(row => row.getCell(2).value === "wallet-refund");
    const cancelledSaleRow = dataRows.find(row => row.getCell(2).value === "sale-cancelled");
    assert.equal(walletRow.getCell(6).value, 100, "wallet refund recorded amount");
    assert.equal(walletRow.getCell(7).value, 0, "wallet refund analytics amount");
    assert.equal(cancelledSaleRow.getCell(6).value, 80, "cancelled sale recorded amount");
    assert.equal(cancelledSaleRow.getCell(7).value, 0, "cancelled sale analytics amount");

    const totalRow = sheet.getRow(sheet.rowCount);
    assert.equal(totalRow.getCell(6).value, 1630, "all recorded payments");
    assert.equal(totalRow.getCell(7).value, 1450, "analytics-recognized payments");
});

test("analytics summary reconciles slots, products, memberships, cancellations, payments, and due", () => {
    const bookings = [
        {
            id: "booking",
            status: "confirmed",
            final_amount: 100,
            product_amount: 20,
            payment: [{ method: "Cash", amount: 120 }],
        },
        {
            id: "membership",
            type: "membership",
            status: "active",
            final_amount: 50,
            payment: [{ method: "UPI", amount: 50 }],
        },
        {
            id: "cancelled-retained",
            status: "cancelled",
            final_amount: 100,
            payment: [{ method: "Online", amount: 30 }],
        },
        {
            id: "cancelled-wallet",
            status: "cancelled",
            wallet_refund_choice: "refund_to_gameon_wallet",
            wallet_refund_amount: 40,
            payment: [{ method: "Online", amount: 40 }],
        },
    ];
    const sales = [
        {
            id: "sale",
            payment_status: "partial",
            total_amount: 60,
            payment: [{ method: "Card", amount: 50 }],
        },
        {
            id: "cancelled-sale",
            payment_status: "cancelled",
            total_amount: 70,
            payment: [{ method: "Cash", amount: 70 }],
        },
    ];

    const summary = calculateAnalyticsSummary(bookings, sales);
    assert.equal(summary.slotRevenue, 100);
    assert.equal(summary.bookingProductRevenue, 20);
    assert.equal(summary.standaloneProductRevenue, 60);
    assert.equal(summary.membershipRevenue, 50);
    assert.equal(summary.retainedCancelledRevenue, 30);
    assert.equal(summary.totalRevenue, 180);
    assert.equal(summary.totalBusinessRevenue, 260);
    assert.equal(summary.totalRecognizedPayments, 250);
    assert.equal(summary.pendingAmount, 10);
});
