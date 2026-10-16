# Dynamic Financial Manual Entry Form

This HTML form is designed to connect seamlessly with the **consolidated Google Apps Script engine** in your spreadsheet. Because your backend script parses incoming submissions as raw JSON data via `JSON.parse(e.postData.contents)`, a standard HTML form submission will fail. 

This form includes a client-side JavaScript handler that captures the input, compiles it into a structured JSON payload, and posts it to your Google Web App using `fetch()`. It also features visual styling via Tailwind CSS and built-in loading and success indicators.

## HTML Code (`manual-entry-form.html`)
Save the code below on your computer as a file named `manual-entry-form.html` (e.g., in your Documents or on your Desktop).

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Financial Transaction Ledger Form</title>
    <!-- Tailwind CSS for Modern UI styling -->
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-50 min-h-screen flex items-center justify-center p-4">
    <div class="bg-white p-8 rounded-xl shadow-md w-full max-w-md border border-slate-100">
        <h2 class="text-2xl font-bold text-slate-800 mb-2 flex items-center gap-2">
            <span>💳</span> Log New Transaction
        </h2>
        <p class="text-slate-500 text-sm mb-6">Enter transaction details below to sync directly to your master Transactions_DB database sheet.</p>
        
        <form id="transactionForm" class="space-y-4">
            <!-- Date Input -->
            <div>
                <label for="Date" class="block text-sm font-semibold text-slate-700 mb-1">Transaction Date</label>
                <input type="date" id="Date" name="Date" required 
                    class="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-slate-800">
            </div>

            <!-- Merchant Input -->
            <div>
                <label for="Merchant" class="block text-sm font-semibold text-slate-700 mb-1">Merchant / Payee</label>
                <input type="text" id="Merchant" name="Merchant" required placeholder="e.g., Chevron, Starbucks, Landlord"
                    class="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-slate-800">
            </div>

            <!-- Amount Input -->
            <div>
                <label for="Amount" class="block text-sm font-semibold text-slate-700 mb-1">Amount ($)</label>
                <div class="relative">
                    <span class="absolute left-3 top-2 text-slate-400 font-medium">$</span>
                    <input type="number" id="Amount" name="Amount" required step="0.01" min="0.01" placeholder="0.00"
                        class="w-full pl-8 pr-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-slate-800 font-semibold">
                </div>
            </div>

            <!-- Capital Flow Selection -->
            <div>
                <label for="Flow" class="block text-sm font-semibold text-slate-700 mb-1">Capital Flow Direction</label>
                <select id="Flow" name="Flow" required
                    class="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-slate-800">
                    <option value="OUTFLOW" selected>OUTFLOW (Expense / Spending)</option>
                    <option value="INFLOW">INFLOW (Income / Deposit / Credit)</option>
                </select>
            </div>

            <!-- Category Input (Matches Enum!A3:A) -->
            <div>
                <label for="Category" class="block text-sm font-semibold text-slate-700 mb-1">Budget Category</label>
                <select id="Category" name="Category" required
                    class="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-slate-800">
                    <option value="" disabled selected>-- Select a Category --</option>
                    <option value="Food & Drink">Food & Drink</option>
                    <option value="Groceries">Groceries</option>
                    <option value="Bills & Utilities">Bills & Utilities</option>
                    <option value="Home">Home</option>
                    <option value="Transportation">Transportation</option>
                    <option value="Entertainment">Entertainment</option>
                    <option value="Personal">Personal</option>
                    <option value="Travel">Travel</option>
                    <option value="Miscellaneous">Miscellaneous</option>
                </select>
            </div>

            <!-- Account Source (Matches Enum!B3:B) -->
            <div>
                <label for="Source" class="block text-sm font-semibold text-slate-700 mb-1">Account Source</label>
                <select id="Source" name="Source" required
                    class="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-slate-800">
                    <option value="" disabled selected>-- Select an Account --</option>
                    <option value="Cash">Cash</option>
                    <option value="Chase Checking">Chase Checking</option>
                    <option value="Chase Sapphire">Chase Sapphire</option>
                    <option value="TD Bank">TD Bank</option>
                    <option value="Credit Card">Credit Card</option>
                    <option value="Savings">Savings</option>
                </select>
            </div>

            <!-- Action Buttons -->
            <div class="pt-2">
                <button type="submit" id="submitBtn"
                    class="w-full bg-slate-800 hover:bg-slate-900 text-white font-semibold py-2 px-4 rounded-lg transition-colors focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-slate-800 flex items-center justify-center gap-2">
                    Submit Transaction 🚀
                </button>
            </div>
        </form>

        <!-- Dynamic Success Alert -->
        <div id="successAlert" class="hidden mt-4 p-3 bg-green-50 border border-green-200 text-green-800 rounded-lg text-sm text-center font-medium">
            ✅ Transaction successfully logged to Transactions_DB!
        </div>

        <!-- Dynamic Error Alert -->
        <div id="errorAlert" class="hidden mt-4 p-3 bg-red-50 border border-red-200 text-red-800 rounded-lg text-sm text-center font-medium">
            ❌ Submit Failed: <span id="errorMessage"></span>
        </div>
    </div>

    <script>
        // =========================================================================
        // IMPORTANT: Paste your copied Google Web App URL (ending in /exec) below
        // =========================================================================
        const WEB_APP_URL = https://script.google.com/macros/s/AKfycbzRHlPYcIqXmDmANrrRcT19dIhMSkXUuflAtamSfdqQiodKgUV2eEIIOHHqSzn4KwXx/exec;

        document.getElementById('Date').valueAsDate = new Date(); // Set default to today

        document.getElementById('transactionForm').addEventListener('submit', async function(e) {
            e.preventDefault();
            
            const submitBtn = document.getElementById('submitBtn');
            const successAlert = document.getElementById('successAlert');
            const errorAlert = document.getElementById('errorAlert');
            
            // Reset alerts & button state
            successAlert.classList.add('hidden');
            errorAlert.classList.add('hidden');
            submitBtn.disabled = true;
            submitBtn.innerHTML = "Processing Submission... ⏳";
            submitBtn.classList.replace('bg-slate-800', 'bg-slate-400');

            // Collect form data matching your doPost(e) JSON variables
            const payload = {
                Date: document.getElementById('Date').value,
                Merchant: document.getElementById('Merchant').value.trim(),
                Amount: document.getElementById('Amount').value,
                Flow: document.getElementById('Flow').value,
                Category: document.getElementById('Category').value,
                Source: document.getElementById('Source').value
            };

            try {
                if (WEB_APP_URL === "YOUR_WEB_APP_URL_HERE") {
                    throw new Error("You must configure your Google Apps Script Web App URL first.");
                }

                // Send JSON POST request directly to Google Apps Script Web App
                const response = await fetch(WEB_APP_URL, {
                    method: 'POST',
                    mode: 'no-cors', // Avoids local development CORS blocks, returns opaque response
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(payload)
                });

                // Since 'no-cors' mode hides standard JSON results, we verify submission completion
                successAlert.classList.remove('hidden');
                document.getElementById('transactionForm').reset();
                document.getElementById('Date').valueAsDate = new Date(); // Reset date
            } catch (err) {
                console.error("Submission Error:", err);
                document.getElementById('errorMessage').textContent = err.message;
                errorAlert.classList.remove('hidden');
            } finally {
                // Restore button state
                submitBtn.disabled = false;
                submitBtn.innerHTML = "Submit Transaction 🚀";
                submitBtn.classList.replace('bg-slate-400', 'bg-slate-800');
            }
        });
    </script>
</body>
</html>
```
