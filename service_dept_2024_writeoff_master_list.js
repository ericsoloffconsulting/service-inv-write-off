/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 * @NAmdConfig /SuiteScripts/ericsoloffconsulting/JsLibraryConfig.json
 * 
 * Service Department 2024 Write-Off Master List
 * 
 * Purpose: Displays all open service-related invoices and credit memos from 2024 and earlier
 * 
 * Includes transactions that either:
 * - Have line-level department 13 (Service)
 * - Customer category is "Service Vendor" (2) or "Old Vendor" (4)
 */
define(['N/ui/serverWidget', 'N/query', 'N/log', 'N/runtime', 'N/url'],
    /**
     * @param {serverWidget} serverWidget
     * @param {query} query
     * @param {log} log
     * @param {runtime} runtime
     * @param {url} url
     */
    function (serverWidget, query, log, runtime, url) {

        /**
         * Handles GET and POST requests to the Suitelet
         * @param {Object} context - NetSuite context object containing request/response
         */
        function onRequest(context) {
            if (context.request.method === 'GET') {
                handleGet(context);
            } else {
                handleGet(context);
            }
        }

        /**
         * Handles GET requests
         * @param {Object} context
         */
        function handleGet(context) {
            var request = context.request;
            var response = context.response;

            log.debug('GET Request', 'Parameters: ' + JSON.stringify(request.parameters));

            var form = serverWidget.createForm({
                title: 'Service Department A/R Transactions'
            });

            try {
                var htmlContent = buildPageHTML(request.parameters);

                var htmlField = form.addField({
                    id: 'custpage_html_content',
                    type: serverWidget.FieldType.INLINEHTML,
                    label: 'Content'
                });

                htmlField.defaultValue = htmlContent;

            } catch (e) {
                log.error('Error Building Form', {
                    error: e.message,
                    stack: e.stack
                });

                var errorField = form.addField({
                    id: 'custpage_error',
                    type: serverWidget.FieldType.INLINEHTML,
                    label: 'Error'
                });
                errorField.defaultValue = '<p style="color:red;">Error loading portal: ' + escapeHtml(e.message) + '</p>';
            }

            context.response.writePage(form);
        }

        /**
         * Builds the main page HTML
         * @param {Object} params - URL parameters
         * @returns {string} HTML content
         */
        function buildPageHTML(params) {
            try {
                log.debug('buildPageHTML Start', 'Params: ' + JSON.stringify(params));
                
                var scriptUrl = url.resolveScript({
                    scriptId: runtime.getCurrentScript().id,
                    deploymentId: runtime.getCurrentScript().deploymentId,
                    returnExternalUrl: false
                });

                var balanceAsOf = (params.balanceAsOf && params.balanceAsOf.trim()) ? params.balanceAsOf.trim() : '2024-12-31';
                log.debug('Balance As Of', balanceAsOf);

                var html = '';

                html += '<div id="loadingOverlay" style="position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(255,255,255,0.95);display:flex;flex-direction:column;justify-content:center;align-items:center;z-index:9999;display:none;">';
                html += '<div style="width:50px;height:50px;border:4px solid #e0e0e0;border-top:4px solid #4CAF50;border-radius:50%;animation:spin 1s linear infinite;"></div>';
                html += '<div style="margin-top:15px;font-size:16px;color:#333;font-weight:500;">Loading data...</div>';
                html += '<style>@keyframes spin{0%{transform:rotate(0deg)}100%{transform:rotate(360deg)}}</style>';
                html += '</div>';

                log.debug('Data Loading', 'Starting to load transaction data for ' + balanceAsOf);
                var transactionResult = searchServiceTransactions(balanceAsOf);
                var transactions = transactionResult.transactions;
                var isTruncated = transactionResult.isTruncated;
                var actualCount = transactionResult.actualCount;
                log.debug('Transactions Loaded', 'Count: ' + transactions.length);

                var totalInvoices = 0;
                var totalInvoiceAmount = 0;
                var totalCredits = 0;
                var totalCreditAmount = 0;
                var totalNetAmount = 0;

                if (isTruncated) {
                    totalInvoices = transactionResult.invoiceCount;
                    totalInvoiceAmount = transactionResult.invoiceTotal;
                    totalCredits = transactionResult.creditCount;
                    totalCreditAmount = transactionResult.creditTotal;
                    totalNetAmount = transactionResult.netTotal;
                } else {
                    for (var i = 0; i < transactions.length; i++) {
                        var txn = transactions[i];
                        var amount = parseFloat(txn.amount_remaining) || 0;
                        totalNetAmount += amount;
                        
                        if (amount > 0) {
                            totalInvoices++;
                            totalInvoiceAmount += amount;
                        } else if (amount < 0) {
                            totalCredits++;
                            totalCreditAmount += Math.abs(amount);
                        }
                    }
                }

                html += '<style>' + getStyles() + '</style>';

                html += '<div class="portal-container">';

                html += '<div class="balance-as-of-section">';
                html += '<label class="balance-as-of-label">Transactions As Of:</label>';
                html += '<input type="date" id="balanceAsOfDate" class="balance-as-of-input" value="' + balanceAsOf + '">';
                html += '<button type="button" id="loadResultsBtn" class="load-results-btn">Load Results</button>';
                html += '</div>';

                html += '<div class="summary-section">';
                html += '<h2 class="summary-title">Open Service Transactions Summary</h2>';
                html += '<div class="summary-grid">';
                html += buildSummaryCard('Open Invoices', totalInvoices, totalInvoiceAmount);
                html += buildSummaryCard('Open Credit Memos', totalCredits, totalCreditAmount);
                html += buildSummaryCard('Net Amount', totalInvoices + totalCredits, totalNetAmount);
                html += '</div>';
                html += '</div>';

                html += buildDataSection('transactions', 'Open Service Transactions', 
                    'Invoices and credit memos from ' + balanceAsOf + ' or earlier with service department criteria', 
                    transactions, scriptUrl, isTruncated, actualCount);

                html += '</div>';

                html += '<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>';

                html += '<script>' + getJavaScript(scriptUrl) + '</script>';

                log.debug('buildPageHTML Complete', 'HTML generated successfully');
                return html;
            
            } catch (e) {
                log.error('Error in buildPageHTML', {
                    error: e.message,
                    stack: e.stack,
                    params: JSON.stringify(params)
                });
                var errorHtml = '<style>body{font-family:Arial,sans-serif;padding:40px;}</style>';
                errorHtml += '<div style="max-width:600px;margin:0 auto;">';
                errorHtml += '<h1 style="color:#d32f2f;">Error Loading Portal</h1>';
                errorHtml += '<p><strong>Error:</strong> ' + escapeHtml(e.message) + '</p>';
                errorHtml += '<p><strong>Details:</strong> Check execution log for details.</p>';
                errorHtml += '<pre style="background:#f5f5f5;padding:15px;overflow:auto;">' + escapeHtml(e.stack || 'No stack trace available') + '</pre>';
                errorHtml += '</div>';
                return errorHtml;
            }
        }

        /**
         * Builds a summary card
         * @param {string} title - Card title
         * @param {number} count - Number of records
         * @param {number} amount - Total amount
         * @returns {string} HTML for summary card
         */
        function buildSummaryCard(title, count, amount) {
            var html = '';
            html += '<div class="summary-card">';
            html += '<div class="summary-card-title">' + escapeHtml(title) + '</div>';
            html += '<div class="summary-card-count">' + count + ' record' + (count !== 1 ? 's' : '') + '</div>';
            html += '<div class="summary-card-amount">' + formatCurrency(amount) + '</div>';
            html += '</div>';
            return html;
        }

        /**
         * Builds a collapsible data section
         * @param {string} sectionId - Section identifier
         * @param {string} title - Section title
         * @param {string} description - Section description
         * @param {Array} data - Data array
         * @param {string} scriptUrl - Suitelet URL
         * @param {boolean} isTruncated - Whether results are truncated
         * @param {number} actualCount - Actual count if truncated
         * @returns {string} HTML for data section
         */
        function buildDataSection(sectionId, title, description, data, scriptUrl, isTruncated, actualCount) {
            var displayedRecords = data.length;
            var totalRecords = isTruncated ? actualCount : displayedRecords;
            var countDisplay = isTruncated 
                ? 'Displaying ' + displayedRecords.toLocaleString() + ' of ' + totalRecords.toLocaleString() + ' records'
                : totalRecords.toLocaleString();
            
            var html = '';
            html += '<div class="search-section" id="section-' + sectionId + '">';
            html += '<div class="search-title collapsible" data-section-id="' + sectionId + '">';
            html += '<span>' + escapeHtml(title) + ' (' + countDisplay + ')' + (isTruncated ? ' <span style="color: #4CAF50; font-size: 11px;">⚠ Totals calculated from all records</span>' : '') + '</span>';
            html += '<span class="toggle-icon" id="toggle-' + sectionId + '">−</span>';
            html += '</div>';
            html += '<div class="search-content" id="content-' + sectionId + '">';
            html += '<div class="search-count">' + escapeHtml(description) + '</div>';
            
            if (totalRecords === 0) {
                html += '<p class="no-results">No open service transactions found.</p>';
            } else {
                html += '<div class="search-box-container">';
                html += '<input type="text" id="searchBox-' + sectionId + '" class="search-box" placeholder="Search this table..." onkeyup="filterTable(\'' + sectionId + '\')">';
                html += '<button type="button" class="export-btn" onclick="exportToExcel(\'' + sectionId + '\')">📥 Export to Excel</button>';
                html += '<span class="search-results-count" id="searchCount-' + sectionId + '"></span>';
                html += '</div>';
                html += buildTransactionTable(data, scriptUrl, sectionId);
            }
            
            html += '</div>';
            html += '</div>';
            return html;
        }

        /**
         * Builds a collapsible summary data section (grouped by customer)
         * @param {string} sectionId - Section identifier
         * @param {string} title - Section title
         * @param {string} description - Section description
         * @param {Array} data - Summary data array
         * @param {string} scriptUrl - Suitelet URL
         * @returns {string} HTML for data section
         */
        function buildSummaryDataSection(sectionId, title, description, data, scriptUrl) {
            var totalCustomers = data.length;
            
            var html = '';
            html += '<div class="search-section" id="section-' + sectionId + '">';
            html += '<div class="search-title collapsible" data-section-id="' + sectionId + '">';
            html += '<span>' + escapeHtml(title) + ' (' + totalCustomers + ' customers)</span>';
            html += '<span class="toggle-icon" id="toggle-' + sectionId + '">−</span>';
            html += '</div>';
            html += '<div class="search-content" id="content-' + sectionId + '">';
            html += '<div class="search-count">' + escapeHtml(description) + '</div>';
            
            if (totalCustomers === 0) {
                html += '<p class="no-results">No customers found.</p>';
            } else {
                html += '<div class="search-box-container">';
                html += '<input type="text" id="searchBox-' + sectionId + '" class="search-box" placeholder="Search this table..." onkeyup="filterTable(\'' + sectionId + '\')" style="flex:1;">';
                html += '<div style="display: flex; align-items: center; margin-left: 10px; border: 1px solid #d3d3d3; border-radius: 4px; overflow: hidden; background: #f5f5f5;">';
                html += '<button type="button" id="detailsShow" onclick="showDetailView()" style="padding: 6px 12px; border: none; background: #2e5fa3; color: white; cursor: pointer; font-size: 11px; font-weight: bold; white-space: nowrap; transition: all 0.2s;">Show Details</button>';
                html += '<button type="button" id="detailsHide" onclick="hideDetailView()" style="padding: 6px 12px; border: none; background: #f5f5f5; color: #333; cursor: pointer; font-size: 11px; white-space: nowrap; transition: all 0.2s; display: none;">Hide Details</button>';
                html += '</div>';
                html += '<button type="button" class="export-btn" onclick="exportToExcel(\'' + sectionId + '\')">📥 Export to Excel</button>';
                html += '<span class="search-results-count" id="searchCount-' + sectionId + '" style="margin-left: 10px;"></span>';
                html += '</div>';
                html += buildSummaryTable(data, scriptUrl, sectionId);
            }
            
            html += '</div>';
            html += '</div>';
            return html;
        }

        /**
         * Builds the customer summary table
         * @param {Array} summary - Summary data
         * @param {string} scriptUrl - Suitelet URL
         * @param {string} sectionId - Section identifier
         * @returns {string} HTML table
         */
        function buildSummaryTable(summary, scriptUrl, sectionId) {
            var html = '';

            html += '<div class="table-container">';
            html += '<table class="data-table" id="table-' + sectionId + '">';
            html += '<thead>';
            html += '<tr>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 0)">Customer</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 1)">Open Invoices</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 2)">Open Credit Memos</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 3)">Net Amount Remaining</th>';
            html += '</tr>';
            html += '</thead>';
            html += '<tbody>';

            for (var i = 0; i < summary.length; i++) {
                var row = summary[i];
                var rowClass = (i % 2 === 0) ? 'even-row' : 'odd-row';
                var netAmount = parseFloat(row.net_amount) || 0;
                var invoiceCount = parseInt(row.invoice_count) || 0;
                var creditCount = parseInt(row.credit_count) || 0;

                html += '<tr class="' + rowClass + '">';

                html += '<td><a href="/app/common/entity/custjob.nl?id=' + row.customer_id + '" target="_blank">' + escapeHtml(row.customer_name || '-') + '</a></td>';

                html += '<td style="text-align: center;">' + invoiceCount + '</td>';

                html += '<td style="text-align: center;">' + creditCount + '</td>';

                html += '<td class="amount' + (netAmount < 0 ? ' credit-amount' : '') + '">' + formatCurrency(netAmount) + '</td>';

                html += '</tr>';
            }

            html += '</tbody>';
            html += '</table>';
            html += '</div>';

            return html;
        }

        /**
         * Builds the transaction data table
         * @param {Array} transactions - Transaction data
         * @param {string} scriptUrl - Suitelet URL
         * @param {string} sectionId - Section identifier
         * @returns {string} HTML table
         */
        function buildTransactionTable(transactions, scriptUrl, sectionId) {
            var html = '';

            html += '<div class="table-container">';
            html += '<table class="data-table" id="table-' + sectionId + '">';
            html += '<thead>';
            html += '<tr>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 0)">Transaction #</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 1)">External ID</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 2)">Internal ID</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 3)">Date</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 4)">Customer</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 5)">Amount Remaining</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 6)">Status</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 7)">Customer Category</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 8)">Selling Location</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 9)">Service Selling Location</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 10)">Qualified Category</th>';
            html += '</tr>';
            html += '</thead>';
            html += '<tbody>';

            for (var i = 0; i < transactions.length; i++) {
                var txn = transactions[i];
                var rowClass = (i % 2 === 0) ? 'even-row' : 'odd-row';
                var amount = parseFloat(txn.amount_remaining) || 0;
                var isCredit = amount < 0;
                var transactionType = isCredit ? 'custcred' : 'custinvc';
                var transactionTypeText = isCredit ? 'Credit Memo' : 'Invoice';

                html += '<tr class="' + rowClass + '">';

                html += '<td><a href="/app/accounting/transactions/' + transactionType + '.nl?id=' + txn.id + '" target="_blank" title="' + transactionTypeText + '">' + escapeHtml(txn.tranid) + '</a></td>';

                html += '<td>' + escapeHtml(txn.externalid || '-') + '</td>';

                html += '<td>' + escapeHtml(txn.id) + '</td>';

                html += '<td data-date="' + escapeHtml(txn.trandate || '') + '">' + formatDate(txn.trandate) + '</td>';

                html += '<td>' + escapeHtml(txn.customer_name || '-') + '</td>';

                html += '<td class="amount' + (isCredit ? ' credit-amount' : '') + '">' + formatCurrency(amount) + '</td>';

                html += '<td>' + escapeHtml(txn.status_name || '-') + '</td>';

                html += '<td>' + escapeHtml(txn.customer_category || '-') + '</td>';

                html += '<td>' + escapeHtml(txn.department_name || '-') + '</td>';

                html += '<td style="text-align: center;">' + (txn.has_line_dept_13 === 'Y' ? '✓' : '') + '</td>';

                html += '<td style="text-align: center;">' + (txn.has_service_category === 'Y' ? '✓' : '') + '</td';

                html += '</tr>';
            }

            html += '</tbody>';
            
            var totalAmount = 0;
            var totalRecords = transactions.length;
            for (var i = 0; i < transactions.length; i++) {
                totalAmount += parseFloat(transactions[i].amount_remaining) || 0;
            }
            
            html += '<tfoot>';
            html += '<tr>';
            html += '<td colspan="5" class="summary-label">Total (' + totalRecords + ' record' + (totalRecords !== 1 ? 's' : '') + '):</td>';
            html += '<td class="amount">' + formatCurrency(totalAmount) + '</td>';
            html += '<td colspan="5"></td>';
            html += '</tr>';
            html += '</tfoot>';
            
            html += '</table>';
            html += '</div>';

            return html;
        }

        /**
         * Searches for all open service-related transactions
         * @param {string} balanceAsOf - Date to filter transactions (YYYY-MM-DD format)
         * @returns {Object} Object with transactions array and metadata
         */
        function searchServiceTransactions(balanceAsOf) {
            var result = {
                transactions: [],
                isTruncated: false,
                actualCount: 0,
                invoiceCount: 0,
                invoiceTotal: 0,
                creditCount: 0,
                creditTotal: 0,
                netTotal: 0
            };

            try {
                var sql = 'SELECT ' +
                    't.id, ' +
                    't.tranid, ' +
                    't.externalid, ' +
                    't.trandate, ' +
                    'BUILTIN.DF(t.entity) as customer_name, ' +
                    'CASE ' +
                    '    WHEN t.type = \'CustInvc\' THEN t.foreignamountunpaid ' +
                    '    WHEN t.type = \'CustCred\' THEN -1 * tl_main.foreignpaymentamountunused ' +
                    'END as amount_remaining, ' +
                    'BUILTIN.DF(t.status) as status_name, ' +
                    'BUILTIN.DF(c.category) as customer_category, ' +
                    'BUILTIN.DF(tl_main.department) as department_name, ' +
                    'CASE ' +
                    '    WHEN EXISTS ( ' +
                    '        SELECT tl.id ' +
                    '        FROM transactionline tl ' +
                    '        WHERE tl.transaction = t.id ' +
                    '        AND tl.department = 13 ' +
                    '    ) THEN \'Y\' ' +
                    '    ELSE \'N\' ' +
                    'END as has_line_dept_13, ' +
                    'CASE ' +
                    '    WHEN c.category IN (2, 4) THEN \'Y\' ' +
                    '    ELSE \'N\' ' +
                    'END as has_service_category ' +
                    'FROM transaction t ' +
                    'INNER JOIN transactionline tl_main ON t.id = tl_main.transaction AND tl_main.mainline = \'T\' ' +
                    'INNER JOIN customer c ON t.entity = c.id ' +
                    'WHERE t.trandate <= TO_DATE(\'' + balanceAsOf + '\', \'YYYY-MM-DD\') ' +
                    'AND t.status = \'A\' ' +
                    'AND t.type IN (\'CustInvc\', \'CustCred\') ' +
                    'AND ( ' +
                    '    EXISTS ( ' +
                    '        SELECT tl.id ' +
                    '        FROM transactionline tl ' +
                    '        WHERE tl.transaction = t.id ' +
                    '        AND tl.department = 13 ' +
                    '    ) ' +
                    '    OR c.category IN (2, 4) ' +
                    ') ' +
                    'ORDER BY t.trandate, t.tranid, t.id';

                log.debug('Service Transaction Query', sql);

                var countSql = 'SELECT ' +
                    'COUNT(DISTINCT t.id) as total_count, ' +
                    'SUM(CASE WHEN t.type = \'CustInvc\' THEN 1 ELSE 0 END) as invoice_count, ' +
                    'SUM(CASE WHEN t.type = \'CustInvc\' THEN t.foreignamountunpaid ELSE 0 END) as invoice_total, ' +
                    'SUM(CASE WHEN t.type = \'CustCred\' THEN 1 ELSE 0 END) as credit_count, ' +
                    'SUM(CASE WHEN t.type = \'CustCred\' THEN tl_main.foreignpaymentamountunused ELSE 0 END) as credit_total, ' +
                    'SUM(CASE ' +
                    '    WHEN t.type = \'CustInvc\' THEN t.foreignamountunpaid ' +
                    '    WHEN t.type = \'CustCred\' THEN -1 * tl_main.foreignpaymentamountunused ' +
                    '    ELSE 0 ' +
                    'END) as net_total ' +
                    'FROM transaction t ' +
                    'INNER JOIN transactionline tl_main ON t.id = tl_main.transaction AND tl_main.mainline = \'T\' ' +
                    'INNER JOIN customer c ON t.entity = c.id ' +
                    'WHERE t.trandate <= TO_DATE(\'' + balanceAsOf + '\', \'YYYY-MM-DD\') ' +
                    'AND t.status = \'A\' ' +
                    'AND t.type IN (\'CustInvc\', \'CustCred\') ' +
                    'AND ( ' +
                    '    EXISTS ( ' +
                    '        SELECT tl.id ' +
                    '        FROM transactionline tl ' +
                    '        WHERE tl.transaction = t.id ' +
                    '        AND tl.department = 13 ' +
                    '    ) ' +
                    '    OR c.category IN (2, 4) ' +
                    ')';

                var countResults = query.runSuiteQL({ query: countSql }).asMappedResults();
                if (countResults.length > 0) {
                    result.actualCount = parseInt(countResults[0].total_count) || 0;
                    result.invoiceCount = parseInt(countResults[0].invoice_count) || 0;
                    result.invoiceTotal = parseFloat(countResults[0].invoice_total) || 0;
                    result.creditCount = parseInt(countResults[0].credit_count) || 0;
                    result.creditTotal = parseFloat(countResults[0].credit_total) || 0;
                    result.netTotal = parseFloat(countResults[0].net_total) || 0;
                }

                var pagedData = query.runSuiteQLPaged({
                    query: sql,
                    pageSize: 1000
                });

                var pageCount = pagedData.pageRanges.length;
                log.debug('Query Pagination', 'Total pages: ' + pageCount + ', Total records: ' + result.actualCount);

                for (var p = 0; p < pageCount; p++) {
                    var pageData = pagedData.fetch({ index: p });
                    var pageResults = pageData.data.asMappedResults();
                    result.transactions = result.transactions.concat(pageResults);
                }
                
                log.debug('All Pages Fetched', 'Total transactions loaded: ' + result.transactions.length);

                var seenIds = {};
                var deduped = [];
                var dupeCount = 0;
                for (var i = 0; i < result.transactions.length; i++) {
                    var txn = result.transactions[i];
                    if (!seenIds[txn.id]) {
                        seenIds[txn.id] = true;
                        deduped.push(txn);
                    } else {
                        dupeCount++;
                        log.debug('Duplicate Found', 'Transaction ID: ' + txn.id + ' (' + txn.tranid + ') - skipping');
                    }
                }
                
                if (dupeCount > 0) {
                    log.audit('Duplicates Removed', 'Found and removed ' + dupeCount + ' duplicate transaction(s)');
                    result.transactions = deduped;
                }

                log.debug('Service Transactions Found', {
                    displayed: result.transactions.length,
                    actual: result.actualCount,
                    truncated: result.isTruncated
                });

            } catch (e) {
                log.error('Error Searching Service Transactions', {
                    error: e.message,
                    stack: e.stack
                });
            }

            return result;
        }

        /**
         * Searches for service transactions grouped by customer
         * @param {string} balanceAsOf - Date to filter transactions (YYYY-MM-DD format)
         * @returns {Object} Object with summary array
         */
        function searchServiceTransactionsSummary(balanceAsOf) {
            var result = {
                summary: []
            };

            try {
                var sql = 'SELECT ' +
                    't.entity as customer_id, ' +
                    'MAX(BUILTIN.DF(t.entity)) as customer_name, ' +
                    'COUNT(CASE WHEN t.type = \'CustInvc\' THEN 1 END) as invoice_count, ' +
                    'COUNT(CASE WHEN t.type = \'CustCred\' THEN 1 END) as credit_count, ' +
                    'SUM(CASE ' +
                    '    WHEN t.type = \'CustInvc\' THEN t.foreignamountunpaid ' +
                    '    WHEN t.type = \'CustCred\' THEN -1 * tl_main.foreignpaymentamountunused ' +
                    '    ELSE 0 ' +
                    'END) as net_amount ' +
                    'FROM transaction t ' +
                    'INNER JOIN transactionline tl_main ON t.id = tl_main.transaction AND tl_main.mainline = \'T\' ' +
                    'INNER JOIN customer c ON t.entity = c.id ' +
                    'WHERE t.trandate <= TO_DATE(\'' + balanceAsOf + '\', \'YYYY-MM-DD\') ' +
                    'AND t.status LIKE \'%A%\' ' +
                    'AND t.type IN (\'CustInvc\', \'CustCred\') ' +
                    'AND ( ' +
                    '    EXISTS ( ' +
                    '        SELECT tl.id ' +
                    '        FROM transactionline tl ' +
                    '        WHERE tl.transaction = t.id ' +
                    '        AND tl.department = 13 ' +
                    '    ) ' +
                    '    OR c.category IN (2, 4) ' +
                    ') ' +
                    'GROUP BY t.entity ' +
                    'ORDER BY net_amount DESC';

                log.debug('Service Transaction Summary Query', sql);

                result.summary = query.runSuiteQL({ query: sql }).asMappedResults();

                log.debug('Service Transaction Summary Found', {
                    customers: result.summary.length
                });

            } catch (e) {
                log.error('Error Searching Service Transaction Summary', {
                    error: e.message,
                    stack: e.stack
                });
            }

            return result;
        }

        /**
         * Formats a currency value
         * @param {number} value - Currency value
         * @returns {string} Formatted currency
         */
        function formatCurrency(value) {
            if (!value && value !== 0) return '-';
            var prefix = value < 0 ? '-$' : '$';
            return prefix + Math.abs(value).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
        }

        /**
         * Formats a date value
         * @param {string} dateValue - Date string
         * @returns {string} Formatted date
         */
        function formatDate(dateValue) {
            if (!dateValue) return '-';

            try {
                var date = new Date(dateValue);
                var month = date.getMonth() + 1;
                var day = date.getDate();
                var year = date.getFullYear();
                return month + '/' + day + '/' + year;
            } catch (e) {
                return dateValue;
            }
        }

        /**
         * Escapes HTML special characters
         * @param {string} text - Text to escape
         * @returns {string} Escaped text
         */
        function escapeHtml(text) {
            if (!text) return '';
            var map = {
                '&': '&amp;',
                '<': '&lt;',
                '>': '&gt;',
                '"': '&quot;',
                "'": '&#039;'
            };
            return String(text).replace(/[&<>"']/g, function (m) { return map[m]; });
        }

        /**
         * Returns CSS styles for the page
         * @returns {string} CSS content
         */
        function getStyles() {
            return '' +
                '.uir-page-title { display: none !important; }' +
                '.uir-page-title-secondline { border: none !important; margin: 0 !important; padding: 0 !important; }' +
                '.uir-record-type { border: none !important; }' +
                '.bglt { border: none !important; }' +
                '.smalltextnolink { border: none !important; }' +
                '' +
                '.portal-container { margin: 0; padding: 20px; border: none; background: transparent; position: relative; }' +
                '' +
                '.balance-as-of-section { background: linear-gradient(135deg, #1a237e 0%, #3949ab 100%); border-radius: 8px; padding: 12px 20px; margin-bottom: 20px; display: flex; align-items: center; justify-content: center; gap: 12px; box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15); }' +
                '.balance-as-of-label { color: white; font-size: 15px; font-weight: bold; margin: 0; }' +
                '.balance-as-of-input { padding: 6px 10px; border: 2px solid #fff; border-radius: 4px; font-size: 14px; font-weight: 600; color: #1a237e; background: #fff; cursor: pointer; }' +
                '.balance-as-of-input:focus { outline: none; box-shadow: 0 0 0 3px rgba(255, 255, 255, 0.5); }' +
                '.load-results-btn { padding: 6px 16px; border: 2px solid #fff; border-radius: 4px; font-size: 14px; font-weight: 600; color: #1a237e; background: #fff; cursor: pointer; transition: background 0.2s, color 0.2s; }' +
                '.load-results-btn:hover { background: #c5cae9; }' +
                '.load-results-btn:active { background: #9fa8da; }' +
                '' +
                '.initial-message { text-align: center; padding: 60px 20px; font-size: 16px; color: #666; }' +
                '' +
                '.summary-section { background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); border: 2px solid #dee2e6; border-radius: 8px; padding: 20px; margin-bottom: 30px; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1); }' +
                '.summary-title { margin: 0 0 15px 0; font-size: 24px; font-weight: bold; color: #333; text-align: center; }' +
                '.summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; }' +
                '.summary-card { background: white; border: 1px solid #dee2e6; border-radius: 6px; padding: 15px; text-align: center; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08); transition: transform 0.2s, box-shadow 0.2s; }' +
                '.summary-card:hover { transform: translateY(-2px); box-shadow: 0 4px 8px rgba(0, 0, 0, 0.12); }' +
                '.summary-card-title { font-size: 12px; color: #666; margin-bottom: 8px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }' +
                '.summary-card-count { font-size: 14px; color: #333; margin-bottom: 8px; }' +
                '.summary-card-amount { font-size: 18px; font-weight: bold; color: #4CAF50; }' +
                '' +
                '.search-section { margin-bottom: 30px; }' +
                '.search-title { font-size: 16px; font-weight: bold; margin: 25px 0 0 0; color: #333; padding: 15px 10px 15px 10px; border-bottom: 2px solid #4CAF50; cursor: pointer; user-select: none; display: flex; justify-content: space-between; align-items: center; position: -webkit-sticky; position: sticky; top: 0; background: white; z-index: 103; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1); }' +
                '.search-title:hover { background-color: #f8f9fa; }' +
                '.search-title.collapsible { padding-left: 10px; padding-right: 10px; }' +
                '.toggle-icon { font-size: 20px; font-weight: bold; color: #4CAF50; transition: transform 0.3s ease; }' +
                '.search-content { transition: max-height 0.3s ease; }' +
                '.search-content.collapsed { display: none; }' +
                '.search-count { font-style: italic; color: #666; margin: 0; font-size: 12px; padding: 10px 10px; background: white; position: -webkit-sticky; position: sticky; top: 51px; z-index: 102; border-bottom: 1px solid #e9ecef; }' +
                '' +
                '.no-results { text-align: center; color: #999; padding: 40px 20px; font-style: italic; }' +
                '' +
                '.search-box-container { margin: 0; padding: 12px 10px 15px 10px; background: white; position: -webkit-sticky; position: sticky; top: 86px; z-index: 102; border-bottom: 5px solid #4CAF50; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); display: flex; align-items: center; gap: 8px; }' +
                '.search-row { display: flex; gap: 10px; align-items: center; }' +
                '.search-box { flex: 1; padding: 10px 12px; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 14px; box-sizing: border-box; }' +
                '.search-box:focus { outline: none; border-color: #4CAF50; box-shadow: 0 0 0 2px rgba(76, 175, 80, 0.15); }' +
                '.search-results-count { display: none; margin-left: 10px; color: #6c757d; font-size: 13px; font-style: italic; }' +
                '.export-btn { padding: 10px 16px; background: #4CAF50; color: white; border: none; border-radius: 4px; font-size: 14px; font-weight: 600; cursor: pointer; white-space: nowrap; transition: background 0.2s; }' +
                '.export-btn:hover { background: #45a049; }' +
                '.export-btn:active { background: #3d8b40; }' +
                '' +
                '.table-container { overflow: visible; }' +
                '' +
                'table.data-table { border-collapse: separate; border-spacing: 0; width: 100%; margin: 0; margin-top: 0 !important; border-left: 1px solid #ddd; border-right: 1px solid #ddd; border-bottom: 1px solid #ddd; background: white; }' +
                'table.data-table thead th { position: -webkit-sticky; position: sticky; top: 157px; z-index: 101; background-color: #f8f9fa; border: 1px solid #ddd; border-top: none; padding: 10px 8px; text-align: left; vertical-align: top; font-weight: bold; color: #333; font-size: 12px; cursor: pointer; user-select: none; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1); margin-top: 0; }' +
                'table.data-table thead th:hover { background-color: #e9ecef; }' +
                'table.data-table th, table.data-table td { border: 1px solid #ddd; padding: 8px; text-align: left; vertical-align: top; color: #000; }' +
                'table.data-table tbody tr:nth-child(even) td { background-color: #f9f9f9; }' +
                'table.data-table tbody tr:hover td { background-color: #e8f4f8; }' +
                'table.data-table a { color: #0c5460; text-decoration: none; }' +
                'table.data-table a:hover { text-decoration: underline; }' +
                'table.data-table td.amount { text-align: right !important; white-space: nowrap; }' +
                'table.data-table td.credit-amount { color: #d9534f; font-weight: bold; }' +
                'table.data-table tfoot { position: -webkit-sticky; position: sticky; bottom: 0; z-index: 100; }' +
                'table.data-table tfoot td { background-color: #f8f9fa; color: #333; font-weight: bold; padding: 10px 8px; border: 1px solid #ddd; box-shadow: 0 -2px 4px rgba(0, 0, 0, 0.1); }' +
                'table.data-table tfoot td.amount { text-align: right; }' +
                'table.data-table tfoot td.summary-label { text-align: right; font-size: 12px; }' +
                'table.data-table tfoot td.summary-count { text-align: center; font-size: 12px; }';
        }

        /**
         * Returns JavaScript for the page
         * @param {string} scriptUrl - Suitelet URL
         * @returns {string} JavaScript content
         */
        function getJavaScript(scriptUrl) {
            return '' +
                '(function() {' +
                '    document.addEventListener(\'click\', function(e) {' +
                '        var target = e.target.closest(\'.search-title.collapsible\');' +
                '        if (target) {' +
                '            var sectionId = target.getAttribute(\'data-section-id\');' +
                '            if (sectionId) {' +
                '                toggleSection(sectionId);' +
                '            }' +
                '        }' +
                '    });' +
                '})();' +
                '' +
                'function toggleSection(sectionId) {' +
                '    var content = document.getElementById(\'content-\' + sectionId);' +
                '    var icon = document.getElementById(\'toggle-\' + sectionId);' +
                '    if (content && icon) {' +
                '        if (content.classList.contains(\'collapsed\')) {' +
                '            content.classList.remove(\'collapsed\');' +
                '            icon.textContent = String.fromCharCode(8722);' +
                '        } else {' +
                '            content.classList.add(\'collapsed\');' +
                '            icon.textContent = \'+\';' +
                '        }' +
                '    }' +
                '}' +
                '' +
                'function showLoading(message) {' +
                '    var overlay = document.getElementById(\'loadingOverlay\');' +
                '    if (overlay) {' +
                '        var textEl = overlay.querySelector(\'div:last-child\');' +
                '        if (textEl && message) textEl.textContent = message;' +
                '        overlay.style.display = \'flex\';' +
                '    }' +
                '}' +
                '' +
                'function hideLoading() {' +
                '    var overlay = document.getElementById(\'loadingOverlay\');' +
                '    if (overlay) overlay.style.display = \'none\';' +
                '}' +
                '' +
                'document.addEventListener(\'DOMContentLoaded\', function() {' +
                '    hideLoading();' +
                '    ' +
                '    var loadResultsBtn = document.getElementById(\'loadResultsBtn\');' +
                '    if (loadResultsBtn) {' +
                '        loadResultsBtn.addEventListener(\'click\', function() {' +
                '            var balanceAsOfInput = document.getElementById(\'balanceAsOfDate\');' +
                '            var newDate = balanceAsOfInput ? balanceAsOfInput.value : null;' +
                '            if (newDate) {' +
                '                showLoading(\'Loading results for \' + newDate + \'...\');' +
                '                var baseUrl = \'' + scriptUrl + '\';' +
                '                var separator = baseUrl.indexOf(\'?\') > -1 ? \'&\' : \'?\';' +
                '                window.location.href = baseUrl + separator + \'balanceAsOf=\' + newDate;' +
                '            }' +
                '        });' +
                '    }' +
                '});' +
                '' +
                'function sortTable(sectionId, columnIndex) {' +
                '    var table = document.getElementById(\'table-\' + sectionId);' +
                '    var tbody = table.querySelector(\'tbody\');' +
                '    var rows = Array.from(tbody.querySelectorAll(\'tr\'));' +
                '    var currentSort = table.getAttribute(\'data-sort-col\');' +
                '    var currentDir = table.getAttribute(\'data-sort-dir\') || \'asc\';' +
                '    var newDir = (currentSort == columnIndex && currentDir == \'asc\') ? \'desc\' : \'asc\';' +
                '    ' +
                '    var headerCell = table.querySelectorAll(\'th\')[columnIndex];' +
                '    var originalText = headerCell.textContent.replace(/ [▲▼]/g, \'\');' +
                '    headerCell.textContent = \'⏳ Sorting...\';' +
                '    headerCell.style.pointerEvents = \'none\';' +
                '    ' +
                '    setTimeout(function() {' +
                '        rows.sort(function(a, b) {' +
                '            var aCell = a.cells[columnIndex];' +
                '            var bCell = b.cells[columnIndex];' +
                '            var aVal = aCell.getAttribute(\'data-date\') || aCell.textContent.trim();' +
                '            var bVal = bCell.getAttribute(\'data-date\') || bCell.textContent.trim();' +
                '            ' +
                '            if (aCell.classList.contains(\'amount\')) {' +
                '                aVal = parseFloat(aVal.replace(/[^0-9.-]/g, \'\')) || 0;' +
                '                bVal = parseFloat(bVal.replace(/[^0-9.-]/g, \'\')) || 0;' +
                '            } else if (aCell.hasAttribute(\'data-date\')) {' +
                '                var parseDate = function(d) {' +
                '                    if (!d || d === \'-\') return 0;' +
                '                    if (d.indexOf(\'/\') > 0) {' +
                '                        var parts = d.split(\'/\');' +
                '                        return parseInt(parts[2]) * 10000 + parseInt(parts[0]) * 100 + parseInt(parts[1]);' +
                '                    }' +
                '                    return parseInt(d.replace(/-/g, \'\'));' +
                '                };' +
                '                aVal = parseDate(aVal);' +
                '                bVal = parseDate(bVal);' +
                '            } else {' +
                '                aVal = aVal.toLowerCase();' +
                '                bVal = bVal.toLowerCase();' +
                '            }' +
                '            ' +
                '            if (aVal < bVal) return newDir === \'asc\' ? -1 : 1;' +
                '            if (aVal > bVal) return newDir === \'asc\' ? 1 : -1;' +
                '            return 0;' +
                '        });' +
                '        ' +
                '        rows.forEach(function(row) { tbody.appendChild(row); });' +
                '        table.setAttribute(\'data-sort-col\', columnIndex);' +
                '        table.setAttribute(\'data-sort-dir\', newDir);' +
                '        ' +
                '        var allHeaders = table.querySelectorAll(\'th\');' +
                '        for (var i = 0; i < allHeaders.length; i++) {' +
                '            var header = allHeaders[i];' +
                '            if (i == columnIndex) {' +
                '                header.textContent = originalText + (newDir === \'asc\' ? \' ▲\' : \' ▼\');' +
                '            } else {' +
                '                var text = header.textContent.replace(/ [▲▼]/g, \'\').trim();' +
                '                header.textContent = text;' +
                '            }' +
                '        }' +
                '        headerCell.style.pointerEvents = \'\';' +
                '    }, 10);' +
                '}' +
                '' +
                'function filterTable(sectionId) {' +
                '    var input = document.getElementById(\'searchBox-\' + sectionId);' +
                '    var filter = input.value.toUpperCase();' +
                '    var tbody = document.querySelector(\'#table-\' + sectionId + \' tbody\');' +
                '    var rows = tbody.querySelectorAll(\'tr\');' +
                '    var visibleCount = 0;' +
                '    var visibleTotal = 0;' +
                '    ' +
                '    for (var i = 0; i < rows.length; i++) {' +
                '        var row = rows[i];' +
                '        var text = row.textContent || row.innerText;' +
                '        if (text.toUpperCase().indexOf(filter) > -1) {' +
                '            row.style.display = \'\';' +
                '            visibleCount++;' +
                '            var amountCell = row.cells[5];' +
                '            if (amountCell) {' +
                '                var amountText = amountCell.textContent.replace(/[^0-9.-]/g, \'\');' +
                '                visibleTotal += parseFloat(amountText) || 0;' +
                '            }' +
                '        } else {' +
                '            row.style.display = \'none\';' +
                '        }' +
                '    }' +
                '    ' +
                '    var countSpan = document.getElementById(\'searchCount-\' + sectionId);' +
                '    if (filter) {' +
                '        countSpan.textContent = \'Showing \' + visibleCount + \' of \' + rows.length + \' results\';' +
                '        countSpan.style.display = \'inline\';' +
                '    } else {' +
                '        countSpan.style.display = \'none\';' +
                '    }' +
                '    ' +
                '    var table = document.getElementById(\'table-\' + sectionId);' +
                '    var tfoot = table ? table.querySelector(\'tfoot\') : null;' +
                '    if (tfoot) {' +
                '        var footerCells = tfoot.querySelectorAll(\'td\');' +
                '        if (footerCells.length > 0) {' +
                '            var prefix = visibleTotal < 0 ? \'-$\' : \'$\';' +
                '            var formatted = prefix + Math.abs(visibleTotal).toFixed(2).replace(/\\d(?=(\\d{3})+\\.)/g, \'$&,\');' +
                '            footerCells[0].textContent = \'Total (\' + visibleCount + \' record\' + (visibleCount !== 1 ? \'s\' : \'\') + \'):\';' +
                '            if (footerCells[1]) {' +
                '                footerCells[1].textContent = formatted;' +
                '            }' +
                '        }' +
                '    }' +
                '}' +
                '' +
                'function exportToExcel(sectionId) {' +
                '    var table = document.getElementById(\'table-\' + sectionId);' +
                '    if (!table) { alert(\'No data to export\'); return; }' +
                '    ' +
                '    var headers = [];' +
                '    var headerCells = table.querySelectorAll(\'thead th\');' +
                '    for (var i = 0; i < headerCells.length; i++) {' +
                '        headers.push(headerCells[i].textContent.replace(/ [▲▼]/g, \'\').trim());' +
                '    }' +
                '    ' +
                '    var data = [headers];' +
                '    var rows = table.querySelectorAll(\'tbody tr\');' +
                '    for (var i = 0; i < rows.length; i++) {' +
                '        var row = rows[i];' +
                '        if (row.style.display === \'none\') continue;' +
                '        var rowData = [];' +
                '        var cells = row.querySelectorAll(\'td\');' +
                '        for (var j = 0; j < cells.length; j++) {' +
                '            var cell = cells[j];' +
                '            var val = cell.textContent.trim();' +
                '            if (cell.classList.contains(\'amount\')) {' +
                '                val = parseFloat(val.replace(/[\\$,]/g, \'\')) || 0;' +
                '            }' +
                '            rowData.push(val);' +
                '        }' +
                '        data.push(rowData);' +
                '    }' +
                '    ' +
                '    var ws = XLSX.utils.aoa_to_sheet(data);' +
                '    ' +
                '    var range = XLSX.utils.decode_range(ws["!ref"]);' +
                '    for (var R = 1; R <= range.e.r; R++) {' +
                '        var addr = XLSX.utils.encode_cell({r: R, c: 5});' +
                '        if (ws[addr]) { ws[addr].z = "\\\"$\\\"#,##0.00"; }' +
                '    }' +
                '    ' +
                '    var wb = XLSX.utils.book_new();' +
                '    XLSX.utils.book_append_sheet(wb, ws, \'Service Transactions\');' +
                '    ' +
                '    var today = new Date();' +
                '    var dateStr = (today.getMonth()+1) + \'-\' + today.getDate() + \'-\' + today.getFullYear();' +
                '    XLSX.writeFile(wb, \'Service_Dept_2024_WriteOff_\' + dateStr + \'.xlsx\');' +
                '}';
        }

        return {
            onRequest: onRequest
        };
    });
