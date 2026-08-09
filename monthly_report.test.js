const test = require("node:test");
const assert = require("node:assert/strict");
const ExcelJS = require("exceljs");

const {
    createOnlineBookingSheet,
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
    assert.equal(row.getCell(15).value, 1000, "legacy online payment");
    assert.equal(row.getCell(20).value, 1000, "legacy total paid");
    assert.equal(row.getCell(23).value, 100, "legacy commission");
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
    assert.equal(row.getCell(15).value, 0, "recognized online payment");
    assert.equal(row.getCell(19).value, 0, "recognized wallet payment");
    assert.equal(row.getCell(20).value, 0, "recognized total paid");
    assert.equal(row.getCell(23).value, 0, "commission");
});
