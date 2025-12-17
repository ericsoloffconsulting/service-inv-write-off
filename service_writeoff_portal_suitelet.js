/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 * 
 * Service Write-Off Portal
 * 
 * Purpose: Displays Service Department (Dept 13) Sales Orders with unbilled line items.
 * Allows bulk selection and processing via "Bill & Write-Off" action.
 * 
 * User can select multiple SOs via checkboxes and submit for bulk processing.
 * Scheduled script will:
 * 1. Transform SO to Invoice
 * 2. If invoice total > $0, create write-off Credit Memo with item 306698
 * 3. Apply credit memo to invoice
 * 
 * Query Logic: Finds Sales Orders in department 13 with unbilled line items:
 * - Status NOT IN ('H' Closed, 'G' Billed)
 * - Has line items not yet invoiced (NOT EXISTS check)
 * - Aggregates unbilled lines and amounts per SO
 */

define(['N/ui/serverWidget', 'N/query', 'N/log', 'N/url', 'N/record'],
    function(serverWidget, query, log, url, record) {

        /**
         * Handles GET requests to the Suitelet
         */
        function onRequest(context) {
            if (context.request.method === 'GET') {
                handleGet(context);
            } else if (context.request.method === 'POST') {
                handlePost(context);
            }
        }

        /**
         * Handles POST requests - queues Sales Orders for write-off processing
         */
        function handlePost(context) {
            var response = context.response;
            var params = context.request.parameters;
            
            try {
                // Get selected SO IDs from checkbox selections
                var selectedSOIds = params.selectedSOIds;
                
                if (!selectedSOIds) {
                    response.setHeader({ name: 'Content-Type', value: 'application/json' });
                    response.write(JSON.stringify({ 
                        success: false, 
                        message: 'No Sales Orders selected. Please select at least one SO.' 
                    }));
                    return;
                }
                
                // Parse comma-separated SO IDs
                var soIdArray = selectedSOIds.split(',').filter(function(id) { return id.trim(); });
                
                log.audit('Queue for Write-Off', 'Queueing ' + soIdArray.length + ' Sales Orders: ' + soIdArray.join(', '));
                
                // Get today's date
                var todayDate = new Date();
                
                // Update each SO with queue date using record.submitFields (only 4 governance units per record)
                var queuedIds = [];
                var failedIds = [];
                
                for (var i = 0; i < soIdArray.length; i++) {
                    var soId = soIdArray[i];
                    try {
                        record.submitFields({
                            type: record.Type.SALES_ORDER,
                            id: soId,
                            values: {
                                custbody_service_queued_for_write_off: todayDate
                            },
                            options: {
                                enableSourcing: false,
                                ignoreMandatoryFields: true
                            }
                        });
                        queuedIds.push(soId);
                        log.audit('SO Queued', 'Sales Order ' + soId + ' queued for write-off on ' + todayDate.toISOString());
                    } catch (e) {
                        failedIds.push(soId);
                        log.error('Queue Error', 'Failed to queue SO ' + soId + ': ' + e.toString());
                    }
                }
                
                // Return success response with queued IDs for dynamic UI update
                var message = queuedIds.length + ' Sales Order(s) queued for write-off processing.';
                if (failedIds.length > 0) {
                    message += ' Failed to queue: ' + failedIds.join(', ');
                }
                
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ 
                    success: true, 
                    message: message,
                    queuedIds: queuedIds,
                    failedIds: failedIds,
                    count: queuedIds.length
                }));
                
            } catch (e) {
                log.error('Queue for Write-Off Error', e.toString());
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ success: false, message: 'Error: ' + e.toString() }));
            }
        }

        /**
         * Handles GET requests - builds and displays the report
         */
        function handleGet(context) {
            var response = context.response;
            var params = context.request.parameters;

            // Check if this is an AJAX request to load data
            if (params.loadData === 'true') {
                return handleLoadData(context);
            }

            log.audit('Service Write-Off Portal', 'Showing initial empty page');

            try {
                // Create NetSuite form (preserves menu bar and navigation)
                var form = serverWidget.createForm({
                    title: 'Service Write-Off Portal'
                });
                
                // Add inline HTML field for custom content
                var htmlField = form.addField({
                    id: 'custpage_html_content',
                    type: serverWidget.FieldType.INLINEHTML,
                    label: ' '
                });
                
                // Build HTML content (without full page structure)
                var html = buildReportHTML(null);
                htmlField.defaultValue = html;
                
                // Write the form page (keeps NetSuite chrome)
                response.writePage(form);

            } catch (e) {
                log.error('Report Generation Error', e.toString());
                var errorForm = serverWidget.createForm({ title: 'Error' });
                var errorField = errorForm.addField({ id: 'custpage_error', type: serverWidget.FieldType.INLINEHTML, label: ' ' });
                errorField.defaultValue = '<div style="color:red;padding:20px;"><h2>Error Generating Report</h2><p>' + e.toString() + '</p></div>';
                response.writePage(errorForm);
            }
        }

        /**
         * Handles AJAX request to load report data
         */
        function handleLoadData(context) {
            var response = context.response;
            
            log.audit('Service Write-Off Portal', 'Loading unbilled SO data via AJAX');

            try {
                // Run the main query - returns one row per SO with aggregated unbilled data
                var salesOrders = runMainQuery();
                
                log.audit('Query Complete', 'Found ' + salesOrders.length + ' Sales Orders with unbilled items');

                // Calculate summary stats for ALL data
                var totalSOs = salesOrders.length;
                var totalUnbilledLineCount = 0;
                var totalUnbilledAmount = 0;
                
                for (var i = 0; i < salesOrders.length; i++) {
                    var so = salesOrders[i];
                    totalUnbilledLineCount += parseInt(so.unbilled_line_count || 0);
                    totalUnbilledAmount += parseFloat(so.total_unbilled_amount || 0) * -1;
                }

                // Build table body HTML on server
                var tableBodyHtml = '';
                for (var t = 0; t < salesOrders.length; t++) {
                    tableBodyHtml += buildTableRow(salesOrders[t]);
                }

                // Return JSON response with pre-rendered HTML and summary stats
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ 
                    success: true, 
                    summaryTotal: totalSOs,
                    summaryTotalLines: totalUnbilledLineCount,
                    summaryTotalAmount: totalUnbilledAmount,
                    tableBodyHtml: tableBodyHtml
                }));

            } catch (e) {
                log.error('Load Data Error', e.toString());
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ success: false, message: e.toString() }));
            }
        }

        /**
         * Runs the main SuiteQL query to get Service dept Sales Orders with unbilled line items
         * @returns {Array} Array of sales order records with aggregated unbilled data
         */
        function runMainQuery() {
            var sql = "SELECT " +
                "so.id AS so_id, " +
                "so.tranid AS so_number, " +
                "so.trandate AS so_date, " +
                "so.entity AS customer_id, " +
                "MAX(cust.altname) AS customer_name, " +
                "so.status AS so_status, " +
                "MAX(BUILTIN.DF(so.status)) AS so_status_text, " +
                "so.custbody_f4n_job_id AS job_id, " +
                "so.custbody_service_queued_for_write_off AS queued_date, " +
                "MAX(BUILTIN.DF(so.custbody_bas_fa_warranty_type)) AS warranty_type, " +
                "so.custbody24 AS epic_auth, " +
                "so.shipdate AS ship_date, " +
                "so.custbody_bas_estimated_ship_date AS est_ship_date, " +
                "so.custbody_f4n_details AS job_details, " +
                "so.custbody21 AS billing_completed_by, " +
                "so.custbody_f4n_job_state AS job_state, " +
                "so.custbody_f4n_scheduled AS scheduled_date, " +
                "so.custbody_f4n_started AS job_started, " +
                "so.custbody_f4n_completed AS job_completed, " +
                "COUNT(so_line.item) AS unbilled_line_count, " +
                "SUM(so_line.netamount) AS total_unbilled_amount, " +
                "LISTAGG(COALESCE(item.itemid, 'Item #' || so_line.item) || '~~' || so_line.quantity || '~~' || so_line.netamount, '||') WITHIN GROUP (ORDER BY so_line.linesequencenumber) AS unbilled_items_detail " +
                "FROM transaction so " +
                "INNER JOIN transactionline so_line ON so_line.transaction = so.id " +
                "LEFT JOIN customer cust ON so.entity = cust.id " +
                "LEFT JOIN item ON so_line.item = item.id " +
                "WHERE so.type = 'SalesOrd' " +
                // Status 'F' = Pending Billing
                // We do not bill the manufacturer until all line items are fulfilled to the customer
                // so the only transactions that are ready for billing are completely fulfilled which 
                // means Pending Billing status. Fulfillment discrepancies must be managed in a separate process.
                "AND so.status = 'F' " +
                "AND so_line.department = 13 " +
                "AND so_line.taxline = 'F' " +
                "AND so_line.mainline = 'F' " +
                "AND so_line.item IS NOT NULL " +
                "AND so_line.quantity != 0 " +
                "AND NOT EXISTS ( " +
                "    SELECT 1 FROM transactionline inv_line " +
                "    WHERE inv_line.createdfrom = so.id " +
                "    AND inv_line.item = so_line.item " +
                "    AND inv_line.taxline = 'F' " +
                "    AND inv_line.mainline = 'F' " +
                ") " +
                "GROUP BY so.id, so.tranid, so.trandate, so.entity, so.status, so.custbody_f4n_job_id, so.custbody_service_queued_for_write_off, so.custbody24, so.shipdate, so.custbody_bas_estimated_ship_date, so.custbody_f4n_details, so.custbody21, so.custbody_f4n_job_state, so.custbody_f4n_scheduled, so.custbody_f4n_started, so.custbody_f4n_completed " +
                "ORDER BY so.tranid";

            log.audit('Running Unbilled SO Query', sql);
            
            var results = query.runSuiteQL({ query: sql }).asMappedResults();
            
            log.audit('Query Results', 'Found ' + results.length + ' Sales Orders with unbilled items');
            
            return results;
        }

        /**
         * Builds the HTML content for the report (embedded in NetSuite form)
         * @param {Array} data - Sales Order data
         * @returns {string} HTML content
         */
        function buildReportHTML(data) {
            // Get current suitelet URL for AJAX calls
            var suiteletUrl = url.resolveScript({
                scriptId: 'customscript_service_writeoff_portal',
                deploymentId: 'customdeploy_service_writeoff_portal'
            });
            
            // If data is null, show empty shell with Load button
            var isInitialLoad = (data === null);
            var displayData = data || [];
            
            // Build embedded HTML - Scripts must come FIRST before any onclick handlers
            var html = '<script>var SUITELET_URL = "' + suiteletUrl + '";</script>' +
                '<script>' + getJavaScript() + '</script>' +
                '<style>' + getStyles() + '</style>' +
                '<div id="loadingOverlay" class="loading-overlay" style="display:none;">' +
                '<div class="loading-spinner"></div>' +
                '<div class="loading-text">Loading write-off data...</div>' +
                '</div>' +
                '<div class="report-container">' +
                '<div id="successMessage" class="success-msg" style="display:none;"></div>' +
                '<div id="loadButtonContainer" class="load-button-container"' + (isInitialLoad ? '' : ' style="display:none;"') + '>' +
                '<button type="button" id="loadDataBtn" class="load-btn" onclick="loadReportData()">🔄 Load Service Write-Off Data</button>' +
                '<p class="load-hint">Click to load Service Department Sales Orders requiring write-off action.</p>' +
                '</div>' +
                '<div id="reportContent"' + (isInitialLoad ? ' class="hidden"' : '') + '>' +
                '<div id="summarySection">' + buildSummarySection(displayData) + '</div>' +
                '<h2 class="section-header">📋 Sales Orders for Write-Off Review</h2>' +
                '<div id="tableSection">' + buildDataTable(displayData) + '</div>' +
                '</div>' +
                '<div id="jobDetailsTooltip" class="job-details-tooltip"><div class="tooltip-header">Job Information</div><div id="jobDetailsContent"></div></div>' +
                '<div id="lineItemsTooltip" class="line-items-tooltip"><div class="tooltip-header">Unbilled Line Items:</div><div id="tooltipContent"></div></div>' +
                '</div>';

            return html;
        }

        /**
         * Builds summary statistics section with dual summaries (all data + selected data)
         * @param {Array} data - Sales Order data
         * @returns {string} Summary HTML
         */
        function buildSummarySection(data) {
            return '<h2 class="section-header">📊 Summary - All Unbilled Sales Orders</h2>' +
                '<div class="summary-grid">' +
                '<div class="summary-card">' +
                '<div class="summary-value" id="summaryTotal">0</div>' +
                '<div class="summary-label">Total Sales Orders</div>' +
                '<div class="summary-sublabel">With Unbilled Line Items</div>' +
                '</div>' +
                '<div class="summary-card card-pending">' +
                '<div class="summary-value" id="summaryTotalLines">0</div>' +
                '<div class="summary-label">Total Unbilled Lines</div>' +
                '<div class="summary-sublabel">Across All SOs</div>' +
                '</div>' +
                '<div class="summary-card card-open">' +
                '<div class="summary-value" id="summaryTotalAmount">$0.00</div>' +
                '<div class="summary-label">Total Unbilled Amount</div>' +
                '<div class="summary-sublabel">All Sales Orders</div>' +
                '</div>' +
                '</div>' +
                '<h2 class="section-header">✅ Summary - Selected Sales Orders</h2>' +
                '<div class="summary-grid">' +
                '<div class="summary-card card-selected">' +
                '<div class="summary-value" id="selectedCount">0</div>' +
                '<div class="summary-label">Selected SOs</div>' +
                '<div class="summary-sublabel">Ready for Processing</div>' +
                '</div>' +
                '<div class="summary-card card-selected-lines">' +
                '<div class="summary-value" id="selectedLines">0</div>' +
                '<div class="summary-label">Selected Unbilled Lines</div>' +
                '<div class="summary-sublabel">From Selected SOs</div>' +
                '</div>' +
                '<div class="summary-card card-selected-amount">' +
                '<div class="summary-value" id="selectedAmount">$0.00</div>' +
                '<div class="summary-label">Selected Unbilled Amount</div>' +
                '<div class="summary-sublabel">From Selected SOs</div>' +
                '</div>' +
                '</div>';
        }

        /**
         * Builds the main data table with checkbox selection
         * @param {Array} data - Sales Order data
         * @returns {string} Table HTML
         */
        function buildDataTable(data) {
            var html = '<div class="table-controls">' +
                '<input type="text" id="searchBox" placeholder="Search table..." onkeyup="filterTable()">' +
                '<button type="button" id="billWriteOffBtn" class="action-btn-large" onclick="handleQueueForWriteOff()" disabled>Queue for Bill & Write-Off Selected</button>' +
                '</div>' +
                '<div class="table-wrapper">' +
                '<table id="dataTable">' +
                '<thead>' +
                '<tr>' +
                '<th class="th-checkbox"><input type="checkbox" id="selectAllCheckbox" onchange="toggleSelectAll(this)"></th>' +
                '<th class="th-slate" onclick="sortTable(1)">Sales Order<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(2)">Queued<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(3)">Job ID<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(4)">Warranty Type<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(5)">EPIC Auth<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(6)">Customer<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(7)">SO Date<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(8)">Ship Date<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(9)">Est Ship Date<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(10)">SO Status<span class="sort-arrow">↕</span></th>' +
                '<th class="th-teal" onclick="sortTable(11)">Unbilled<br>Line Count<span class="sort-arrow">↕</span></th>' +
                '<th class="th-teal" onclick="sortTable(12)">Total Unbilled<br>Amount<span class="sort-arrow">↕</span></th>' +
                '</tr>' +
                '</thead>' +
                '<tbody id="reportTableBody">';

            var dataLength = data ? data.length : 0;
            for (var i = 0; i < dataLength; i++) {
                html += buildTableRow(data[i]);
            }

            html += '</tbody></table></div>';
            return html;
        }

        /**
         * Builds a single table row with checkbox for selection
         * @param {Object} record - Sales Order record data
         * @returns {string} Table row HTML
         */
        function buildTableRow(record) {
            var soId = record.so_id;
            var unbilledLines = parseInt(record.unbilled_line_count || 0);
            var unbilledAmount = parseFloat(record.total_unbilled_amount || 0) * -1;
            var unbilledDetail = record.unbilled_items_detail || '';
            var jobDetails = record.job_details || '';
            var billingCompletedBy = record.billing_completed_by || '';
            var jobState = record.job_state || '';
            var scheduledDate = record.scheduled_date || '';
            var jobStarted = record.job_started || '';
            var jobCompleted = record.job_completed || '';
            var queuedDate = record.queued_date || '';
            
            var html = '<tr class="line-items-row" data-so-id="' + soId + '" data-unbilled-lines="' + unbilledLines + '" data-unbilled-amount="' + unbilledAmount + '" data-unbilled-detail="' + unbilledDetail + '" data-job-details="' + escapeHtml(jobDetails) + '" data-billing-completed-by="' + escapeHtml(billingCompletedBy) + '" data-job-state="' + escapeHtml(jobState) + '" data-scheduled-date="' + scheduledDate + '" data-job-started="' + jobStarted + '" data-job-completed="' + jobCompleted + '" onmouseenter="showLineItemsTooltip(this); showJobDetailsTooltip(this);" onmouseleave="hideLineItemsTooltip(); hideJobDetailsTooltip();">';
            
            // Checkbox column
            html += '<td class="col-checkbox"><input type="checkbox" class="so-checkbox" value="' + soId + '" onchange="updateSelectedSummary()"></td>';
            
            // Sales Order columns (slate group)
            html += '<td class="col-slate">' + buildTransactionLink(record.so_id, record.so_number, 'salesord') + '</td>';
            html += '<td class="col-slate queued-cell" id="queued-cell-' + soId + '">' + (queuedDate ? '✓' : '') + '</td>';
            html += '<td class="col-slate">' + escapeHtml(record.job_id || '') + '</td>';
            html += '<td class="col-slate">' + escapeHtml(record.warranty_type || '') + '</td>';
            html += '<td class="col-slate">' + escapeHtml(record.epic_auth || '') + '</td>';
            html += '<td class="col-slate">' + buildCustomerLink(record.customer_id, record.customer_name) + '</td>';
            html += '<td class="col-slate">' + formatDate(record.so_date) + '</td>';
            html += '<td class="col-slate">' + formatDate(record.ship_date) + '</td>';
            html += '<td class="col-slate">' + formatDate(record.est_ship_date) + '</td>';
            html += '<td class="col-slate">' + escapeHtml((record.so_status_text || '').replace(/^Sales Order\s*:\s*/i, '')) + '</td>';
            
            // Unbilled data (teal group)
            html += '<td class="col-teal amount">' + unbilledLines + '</td>';
            html += '<td class="col-teal amount">';
            html += '$' + formatAmount(unbilledAmount);
            html += '</td>';
            
            html += '</tr>';
            return html;
        }

        /**
         * Builds tooltip HTML for unbilled line items
         * @param {string} itemsDetail - Delimited string of item details (item~~qty~~amount||...)
         * @returns {string} Tooltip HTML
         */
        function buildLineItemsTooltip(itemsDetail) {
            if (!itemsDetail) return '';
            
            var html = '<div class="line-items-tooltip">';
            html += '<div class="tooltip-header">Unbilled Line Items:</div>';
            html += '<table class="tooltip-table">';
            html += '<tr><th>Item</th><th>Qty</th><th>Amount</th></tr>';
            
            // Parse the delimited string: item~~qty~~amount||item~~qty~~amount
            var items = itemsDetail.split('||');
            for (var i = 0; i < items.length; i++) {
                var parts = items[i].split('~~');
                if (parts.length === 3) {
                    var itemName = parts[0];
                    var qty = parseFloat(parts[1] || 0);
                    var amount = parseFloat(parts[2] || 0);
                    
                    html += '<tr>';
                    html += '<td class="tooltip-item">' + escapeHtml(itemName) + '</td>';
                    html += '<td class="tooltip-qty">' + qty + '</td>';
                    html += '<td class="tooltip-amount">$' + formatAmount(amount) + '</td>';
                    html += '</tr>';
                }
            }
            
            html += '</table>';
            html += '</div>';
            return html;
        }

        /**
         * Builds a clickable transaction link
         * @param {string} id - Transaction internal ID
         * @param {string} tranid - Transaction number
         * @param {string} type - Transaction type
         * @returns {string} Link HTML
         */
        function buildTransactionLink(id, tranid, type) {
            if (!id || !tranid) {
                return '<span class="no-data">—</span>';
            }
            
            var recordType = type.toLowerCase();
            var urlPath = '/app/accounting/transactions/' + recordType + '.nl?id=' + id;
            
            return '<a href="' + urlPath + '" target="_blank" class="transaction-link">' +
                   escapeHtml(tranid) + '</a>';
        }

        /**
         * Builds a clickable customer link
         * @param {string} id - Customer internal ID
         * @param {string} name - Customer name
         * @returns {string} Link HTML
         */
        function buildCustomerLink(id, name) {
            if (!id || !name) {
                return '<span class="no-data">—</span>';
            }
            
            var urlPath = '/app/common/entity/custjob.nl?id=' + id;
            return '<a href="' + urlPath + '" target="_blank" class="customer-link">' +
                   escapeHtml(name) + '</a>';
        }

        /**
         * Formats a currency amount
         * @param {number} amount
         * @returns {string} Formatted amount
         */
        function formatAmount(amount) {
            if (!amount && amount !== 0) return '0.00';
            return parseFloat(amount).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
        }

        /**
         * Formats a date value
         * @param {string} dateValue
         * @returns {string} Formatted date
         */
        function formatDate(dateValue) {
            if (!dateValue) return '';
            try {
                var d = new Date(dateValue);
                return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
            } catch (e) {
                return dateValue;
            }
        }

        /**
         * Escapes HTML special characters
         * @param {string} text
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
            return String(text).replace(/[&<>"']/g, function(m) { return map[m]; });
        }

        /**
         * Returns CSS styles for the report
         * @returns {string} CSS content
         */
        function getStyles() {
            return '.report-container { max-width: 1800px; margin: 10px auto; padding: 10px 5px; font-family: Arial, sans-serif; font-size: 13px; }' +
                '.section-header { color: #013220; margin-top: 25px; margin-bottom: 15px; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; font-size: 18px; }' +
                /* Summary cards */
                '.summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; margin: 15px 0 25px 0; }' +
                '.summary-card { background: #E6EEEA; padding: 16px; border-radius: 6px; text-align: center; border-left: 4px solid #013220; }' +
                '.summary-card.card-pending { border-left-color: #014421; background: #E8F2EC; }' +
                '.summary-card.card-open { border-left-color: #355E3B; background: #EBF0EB; }' +
                '.summary-card.card-selected { border-left-color: #8A9A5B; background: #F4F7F0; }' +
                '.summary-card.card-selected-lines { border-left-color: #6B7F3F; background: #F2F5ED; }' +
                '.summary-card.card-selected-amount { border-left-color: #556B2F; background: #F0F3EC; }' +
                '.summary-value { font-size: 28px; font-weight: bold; color: #013220; margin-bottom: 5px; }' +
                '.summary-card.card-pending .summary-value { color: #014421; }' +
                '.summary-card.card-open .summary-value { color: #355E3B; }' +
                '.summary-card.card-selected .summary-value { color: #8A9A5B; }' +
                '.summary-card.card-selected-lines .summary-value { color: #6B7F3F; }' +
                '.summary-card.card-selected-amount .summary-value { color: #556B2F; }' +
                '.summary-label { font-size: 13px; color: #2d3a33; font-weight: 600; }' +
                '.summary-sublabel { font-size: 11px; color: #6b7c72; margin-top: 4px; font-style: italic; }' +
                /* Table controls */
                '.table-controls { margin: 20px 0; display: flex; gap: 10px; align-items: center; }' +
                '#searchBox { flex: 1; padding: 8px 12px; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 13px; }' +
                '#searchBox:focus { outline: none; border-color: #013220; box-shadow: 0 0 0 2px rgba(1,50,32,0.15); }' +
                '.action-btn-large { background: #013220; color: white; border: none; padding: 10px 24px; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 600; transition: background 0.15s; }' +
                '.action-btn-large:hover:not(:disabled) { background: #012618; }' +
                '.action-btn-large:disabled { background: #cbd5e1; cursor: not-allowed; }' +
                /* Table styling */
                '.table-wrapper { margin: 20px 0; }' +
                '#dataTable { border-collapse: separate; border-spacing: 0; width: 100%; font-size: 12px; background: white; border: 1px solid #cbd5e1; }' +
                '#dataTable th { padding: 10px 8px; text-align: left; font-weight: 600; cursor: pointer; user-select: none; color: white; border: none; border-bottom: 2px solid #cbd5e1; position: -webkit-sticky; position: sticky; top: 0; z-index: 100; box-shadow: 0 2px 4px rgba(0,0,0,0.15); }' +
                '.th-checkbox { background: #013220; cursor: default !important; text-align: center; width: 40px; }' +
                '.th-slate { background: #013220; }' +
                '.th-slate:hover { background: #012618; }' +
                '.th-teal { background: #355E3B; }' +
                '.th-teal:hover { background: #2a4a2f; }' +
                '.sort-arrow { margin-left: 5px; opacity: 0.7; font-size: 10px; }' +
                '#dataTable td { border: none; padding: 8px 6px; vertical-align: middle; font-size: 12px; font-family: Arial, sans-serif; color: #1a2e1f; position: relative; z-index: 1; }' +
                '#dataTable tbody tr { border-bottom: 1px solid #d4e0d7; }' +
                '.col-checkbox { background: #E6EBE9; text-align: center; }' +
                '.col-slate { background: #E6EBE9; }' +
                '.col-teal { background: #EBF0EB; }' +
                '#dataTable tbody tr:hover td.col-checkbox { background: #CDD7D3; }' +
                '#dataTable tbody tr:hover td.col-slate { background: #CDD7D3; }' +
                '#dataTable tbody tr:hover td.col-teal { background: #D7E0D8; }' +
                '.amount { text-align: right; white-space: nowrap; color: #1a2e1f; }' +
                '.no-data { color: #6b7c72; font-style: italic; font-size: 12px; }' +
                '.transaction-link { color: #013220; text-decoration: none; font-weight: 600; font-size: 12px; }' +
                '.transaction-link:hover { text-decoration: underline; color: #355E3B; }' +
                '.customer-link { color: #2d3a33; text-decoration: none; font-size: 12px; }' +
                '.customer-link:hover { text-decoration: underline; color: #013220; }' +
                '.queued-cell { text-align: center; font-size: 16px; color: #355E3B; font-weight: bold; }' +
                /* Line items tooltip */
                '.line-items-row { cursor: help; }' +
                '.line-items-tooltip { display: none; position: fixed; bottom: 20px; right: 20px; background: white; border: 2px solid #355E3B; border-radius: 6px; padding: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.25); z-index: 100000; min-width: 300px; max-width: 400px; }' +
                '.line-items-tooltip.visible { display: block; }' +
                '.tooltip-header { font-weight: 600; color: #013220; margin-bottom: 8px; font-size: 14px; border-bottom: 1px solid #cbd5e1; padding-bottom: 4px; }' +
                '.tooltip-table { width: 100%; border-collapse: collapse; font-size: 13px; }' +
                '.tooltip-table th { background: #355E3B; color: white; padding: 4px 8px; text-align: left; font-size: 12px; font-weight: 600; }' +
                '.tooltip-table th:nth-child(2), .tooltip-table th:nth-child(3) { text-align: right; }' +
                '.tooltip-table td { padding: 4px 8px; border-bottom: 1px solid #e2e8f0; color: #1a2e1f; }' +
                '.tooltip-table tr:last-child td { border-bottom: none; }' +
                '.tooltip-table tr:hover { background: #f8faf9; }' +
                '.tooltip-item { font-weight: 500; }' +
                '.tooltip-qty, .tooltip-amount { text-align: right; }' +
                '.tooltip-amount { font-weight: 600; color: #355E3B; }' +
                /* Job details tooltip */
                '.job-details-tooltip { display: none; position: fixed; bottom: 20px; right: 340px; background: white; border: 2px solid #013220; border-radius: 6px; padding: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.25); z-index: 100000; min-width: 300px; max-width: 400px; }' +
                '.job-details-tooltip.visible { display: block; }' +
                '.job-detail-section { margin-bottom: 12px; }' +
                '.job-detail-section:last-child { margin-bottom: 0; }' +
                '.job-detail-label { font-weight: 600; color: #013220; font-size: 11px; text-transform: uppercase; margin-bottom: 4px; }' +
                '.job-detail-value { color: #1a2e1f; font-size: 12px; line-height: 1.4; white-space: pre-wrap; word-wrap: break-word; }' +
                /* Loading overlay */
                '.loading-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(255,255,255,0.95); z-index: 9999; display: flex; flex-direction: column; justify-content: center; align-items: center; }' +
                '.loading-spinner { border: 4px solid #E6EEEA; border-top: 4px solid #013220; border-radius: 50%; width: 50px; height: 50px; animation: spin 1s linear infinite; }' +
                '@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }' +
                '.loading-text { margin-top: 15px; font-size: 16px; color: #013220; }' +
                /* Success message */
                '.success-msg { background-color: #ECF3ED; border: 1px solid #355E3B; border-left: 4px solid #355E3B; border-radius: 6px; padding: 15px 20px; margin-bottom: 20px; color: #1a2e1f; font-size: 14px; }' +
                '.success-msg .success-title { color: #355E3B; display: block; margin-bottom: 8px; font-size: 15px; font-weight: bold; }' +
                /* Load button */
                '.load-button-container { text-align: center; padding: 40px; }' +
                '.load-btn { background: #013220; color: white; border: none; padding: 15px 40px; border-radius: 6px; cursor: pointer; font-size: 16px; font-weight: 600; box-shadow: 0 2px 4px rgba(1,50,32,0.3); transition: all 0.2s; }' +
                '.load-btn:hover { background: #012618; transform: translateY(-1px); box-shadow: 0 4px 8px rgba(1,50,32,0.4); }' +
                '.load-hint { margin-top: 15px; color: #013220; font-size: 13px; }' +
                '.hidden { display: none !important; }';
        }

        /**
         * Returns JavaScript for interactive features
         * @returns {string} JavaScript content
         */
        function getJavaScript() {
            return 'function filterTable() {' +
                '  var input = document.getElementById("searchBox");' +
                '  var filter = input.value.toUpperCase();' +
                '  var tbody = document.getElementById("reportTableBody");' +
                '  var tr = tbody.children;' +
                '  for (var i = 0; i < tr.length; i++) {' +
                '    var row = tr[i];' +
                '    var txtValue = row.textContent || row.innerText;' +
                '    if (txtValue.toUpperCase().indexOf(filter) > -1) {' +
                '      row.style.display = "";' +
                '    } else {' +
                '      row.style.display = "none";' +
                '    }' +
                '  }' +
                '}' +
                'var sortDir = {};' +
                'function sortTable(n) {' +
                '  var table = document.getElementById("dataTable");' +
                '  var tbody = document.getElementById("reportTableBody");' +
                '  var rows = Array.from(tbody.rows);' +
                '  var numericCols = [11, 12];' +
                '  var dateCols = [7, 8, 9];' +
                '  var isNumeric = numericCols.indexOf(n) > -1;' +
                '  var isDate = dateCols.indexOf(n) > -1;' +
                '  sortDir[n] = sortDir[n] === "asc" ? "desc" : "asc";' +
                '  var dir = sortDir[n];' +
                '  rows.sort(function(a, b) {' +
                '    var xCell = a.cells[n];' +
                '    var yCell = b.cells[n];' +
                '    if (!xCell || !yCell) return 0;' +
                '    var xVal = (xCell.innerText || "").split("\\n")[0].trim();' +
                '    var yVal = (yCell.innerText || "").split("\\n")[0].trim();' +
                '    var xCmp, yCmp;' +
                '    if (isNumeric) {' +
                '      xCmp = parseFloat(xVal.replace(/[^0-9.-]/g, "")) || 0;' +
                '      yCmp = parseFloat(yVal.replace(/[^0-9.-]/g, "")) || 0;' +
                '    } else if (isDate) {' +
                '      xCmp = xVal ? new Date(xVal).getTime() : 0;' +
                '      yCmp = yVal ? new Date(yVal).getTime() : 0;' +
                '    } else {' +
                '      xCmp = xVal.toLowerCase();' +
                '      yCmp = yVal.toLowerCase();' +
                '    }' +
                '    if (xCmp < yCmp) return dir === "asc" ? -1 : 1;' +
                '    if (xCmp > yCmp) return dir === "asc" ? 1 : -1;' +
                '    return 0;' +
                '  });' +
                '  rows.forEach(function(row) { tbody.appendChild(row); });' +
                '}' +
                'function toggleSelectAll(checkbox) {' +
                '  var checkboxes = document.querySelectorAll(".so-checkbox");' +
                '  for (var i = 0; i < checkboxes.length; i++) {' +
                '    checkboxes[i].checked = checkbox.checked;' +
                '  }' +
                '  updateSelectedSummary();' +
                '}' +
                'function updateSelectedSummary() {' +
                '  var checkboxes = document.querySelectorAll(".so-checkbox:checked");' +
                '  var selectedCount = checkboxes.length;' +
                '  var selectedLines = 0;' +
                '  var selectedAmount = 0;' +
                '  for (var i = 0; i < checkboxes.length; i++) {' +
                '    var row = checkboxes[i].closest("tr");' +
                '    if (row) {' +
                '      selectedLines += parseInt(row.getAttribute("data-unbilled-lines") || 0);' +
                '      selectedAmount += parseFloat(row.getAttribute("data-unbilled-amount") || 0);' +
                '    }' +
                '  }' +
                '  var selectedCountEl = document.getElementById("selectedCount");' +
                '  var selectedLinesEl = document.getElementById("selectedLines");' +
                '  var selectedAmountEl = document.getElementById("selectedAmount");' +
                '  var billWriteOffBtn = document.getElementById("billWriteOffBtn");' +
                '  if (selectedCountEl) selectedCountEl.textContent = selectedCount;' +
                '  if (selectedLinesEl) selectedLinesEl.textContent = selectedLines;' +
                '  if (selectedAmountEl) selectedAmountEl.textContent = "$" + selectedAmount.toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",");' +
                '  if (billWriteOffBtn) {' +
                '    billWriteOffBtn.disabled = (selectedCount === 0);' +
                '  }' +
                '}' +
                'function handleQueueForWriteOff() {' +
                '  var checkboxes = document.querySelectorAll(".so-checkbox:checked");' +
                '  if (checkboxes.length === 0) {' +
                '    alert("Please select at least one Sales Order.");' +
                '    return;' +
                '  }' +
                '  var soIds = [];' +
                '  for (var i = 0; i < checkboxes.length; i++) {' +
                '    soIds.push(checkboxes[i].value);' +
                '  }' +
                '  if (!confirm("Queue " + soIds.length + " Sales Order(s) for Bill & Write-Off processing?\\n\\nThis will mark them as queued for future processing.")) {' +
                '    return;' +
                '  }' +
                '  showLoading();' +
                '  var xhr = new XMLHttpRequest();' +
                '  xhr.open("POST", SUITELET_URL, true);' +
                '  xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");' +
                '  xhr.onreadystatechange = function() {' +
                '    if (xhr.readyState === 4) {' +
                '      hideLoading();' +
                '      try {' +
                '        var resp = JSON.parse(xhr.responseText);' +
                '        if (resp.success) {' +
                '          for (var i = 0; i < resp.queuedIds.length; i++) {' +
                '            var queuedCell = document.getElementById("queued-cell-" + resp.queuedIds[i]);' +
                '            if (queuedCell) {' +
                '              queuedCell.textContent = "✓";' +
                '            }' +
                '            var checkbox = document.querySelector(".so-checkbox[value=\\"" + resp.queuedIds[i] + "\\"]");' +
                '            if (checkbox) {' +
                '              checkbox.checked = false;' +
                '              checkbox.disabled = true;' +
                '            }' +
                '          }' +
                '          updateSelectedSummary();' +
                '          alert(resp.message);' +
                '        } else {' +
                '          alert("Error: " + resp.message);' +
                '        }' +
                '      } catch (e) {' +
                '        alert("Error processing response: " + e.toString());' +
                '      }' +
                '    }' +
                '  };' +
                '  xhr.send("selectedSOIds=" + soIds.join(","));' +
                '}' +
                'function showLoading() {' +
                '  document.getElementById("loadingOverlay").style.display = "flex";' +
                '}' +
                'function hideLoading() {' +
                '  document.getElementById("loadingOverlay").style.display = "none";' +
                '}' +
                'function loadReportData() {' +
                '  try {' +
                '    document.getElementById("loadButtonContainer").style.display = "none";' +
                '    showLoading();' +
                '    var xhr = new XMLHttpRequest();' +
                '    xhr.open("GET", SUITELET_URL + "&loadData=true", true);' +
                '    xhr.onreadystatechange = function() {' +
                '      if (xhr.readyState === 4) {' +
                '        hideLoading();' +
                '        try {' +
                '          var resp = JSON.parse(xhr.responseText);' +
                '          if (resp.success) {' +
                '            var summaryTotal = document.getElementById("summaryTotal");' +
                '            var summaryTotalLines = document.getElementById("summaryTotalLines");' +
                '            var summaryTotalAmount = document.getElementById("summaryTotalAmount");' +
                '            var reportTableBody = document.getElementById("reportTableBody");' +
                '            var reportContent = document.getElementById("reportContent");' +
                '            if (summaryTotal) summaryTotal.textContent = resp.summaryTotal;' +
                '            if (summaryTotalLines) summaryTotalLines.textContent = resp.summaryTotalLines;' +
                '            if (summaryTotalAmount) summaryTotalAmount.textContent = "$" + parseFloat(resp.summaryTotalAmount).toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",");' +
                '            if (reportTableBody) reportTableBody.innerHTML = resp.tableBodyHtml;' +
                '            if (reportContent) {' +
                '              reportContent.className = "";' +
                '              reportContent.style.display = "block";' +
                '            }' +
                '          } else {' +
                '            alert("Error loading data: " + resp.message);' +
                '            var loadBtn = document.getElementById("loadButtonContainer");' +
                '            if (loadBtn) loadBtn.style.display = "block";' +
                '          }' +
                '        } catch (e) {' +
                '          alert("Error parsing response: " + e.toString());' +
                '          var loadBtn = document.getElementById("loadButtonContainer");' +
                '          if (loadBtn) loadBtn.style.display = "block";' +
                '        }' +
                '      }' +
                '    };' +
                '    xhr.send();' +
                '  } catch (e) {' +
                '    alert("Error in loadReportData: " + e.toString());' +
                '  }' +
                '}' +
                'function showLineItemsTooltip(row) {' +
                '  var tooltip = document.getElementById("lineItemsTooltip");' +
                '  var tooltipContent = document.getElementById("tooltipContent");' +
                '  var tooltipHeader = tooltip ? tooltip.querySelector(".tooltip-header") : null;' +
                '  if (!tooltip || !tooltipContent) return;' +
                '  var itemsDetail = row.getAttribute("data-unbilled-detail");' +
                '  if (!itemsDetail) return;' +
                '  var totalAmount = parseFloat(row.getAttribute("data-unbilled-amount") || 0);' +
                '  var formattedTotal = "$" + totalAmount.toFixed(2).replace(/\\\\B(?=(\\\\d{3})+(?!\\\\d))/g, ",");' +
                '  if (tooltipHeader) tooltipHeader.textContent = "Unbilled Line Items: " + formattedTotal;' +
                '  var html = "<table class=\\"tooltip-table\\"><tr><th>Item</th><th>Qty</th><th>Amount</th></tr>";' +
                '  var items = itemsDetail.split("||");' +
                '  for (var i = 0; i < items.length; i++) {' +
                '    var parts = items[i].split("~~");' +
                '    if (parts.length === 3) {' +
                '      var itemName = parts[0];' +
                '      var qty = parseFloat(parts[1] || 0) * -1;' +
                '      var amount = parseFloat(parts[2] || 0) * -1;' +
                '      html += "<tr><td class=\\"tooltip-item\\">" + itemName + "</td><td class=\\"tooltip-qty\\">" + qty + "</td><td class=\\"tooltip-amount\\">$" + amount.toFixed(2).replace(/\\\\B(?=(\\\\d{3})+(?!\\\\d))/g, ",") + "</td></tr>";' +
                '    }' +
                '  }' +
                '  html += "</table>";' +
                '  tooltipContent.innerHTML = html;' +
                '  tooltip.className = "line-items-tooltip visible";' +
                '}' +
                'function hideLineItemsTooltip() {' +
                '  var tooltip = document.getElementById("lineItemsTooltip");' +
                '  if (tooltip) tooltip.className = "line-items-tooltip";' +
                '}' +
                'function showJobDetailsTooltip(row) {' +
                '  var tooltip = document.getElementById("jobDetailsTooltip");' +
                '  var tooltipContent = document.getElementById("jobDetailsContent");' +
                '  if (!tooltip || !tooltipContent) return;' +
                '  var jobDetails = row.getAttribute("data-job-details");' +
                '  var billingCompletedBy = row.getAttribute("data-billing-completed-by");' +
                '  var jobState = row.getAttribute("data-job-state");' +
                '  var scheduledDate = row.getAttribute("data-scheduled-date");' +
                '  var jobStarted = row.getAttribute("data-job-started");' +
                '  var jobCompleted = row.getAttribute("data-job-completed");' +
                '  var html = "";' +
                '  if (jobDetails) {' +
                '    var unescapedDetails = jobDetails.replace(/&lt;br&gt;/gi, "<br>");' +
                '    html += "<div class=\\"job-detail-section\\"><div class=\\"job-detail-label\\">Job Details:</div><div class=\\"job-detail-value\\">" + unescapedDetails + "</div></div>";' +
                '  }' +
                '  if (billingCompletedBy) {' +
                '    html += "<div class=\\"job-detail-section\\"><div class=\\"job-detail-label\\">Billing Completed By:</div><div class=\\"job-detail-value\\">" + billingCompletedBy + "</div></div>";' +
                '  }' +
                '  if (jobState) {' +
                '    html += "<div class=\\"job-detail-section\\"><div class=\\"job-detail-label\\">Job State:</div><div class=\\"job-detail-value\\">" + jobState + "</div></div>";' +
                '  }' +
                '  if (scheduledDate) {' +
                '    var formattedScheduled = new Date(scheduledDate).toLocaleDateString();' +
                '    html += "<div class=\\"job-detail-section\\"><div class=\\"job-detail-label\\">Scheduled Date:</div><div class=\\"job-detail-value\\">" + formattedScheduled + "</div></div>";' +
                '  }' +
                '  if (jobStarted) {' +
                '    var formattedStarted = new Date(jobStarted).toLocaleDateString();' +
                '    html += "<div class=\\"job-detail-section\\"><div class=\\"job-detail-label\\">Job Started:</div><div class=\\"job-detail-value\\">" + formattedStarted + "</div></div>";' +
                '  }' +
                '  if (jobCompleted) {' +
                '    var formattedCompleted = new Date(jobCompleted).toLocaleDateString();' +
                '    html += "<div class=\\"job-detail-section\\"><div class=\\"job-detail-label\\">Job Completed:</div><div class=\\"job-detail-value\\">" + formattedCompleted + "</div></div>";' +
                '  }' +
                '  if (!html) {' +
                '    html = "<div class=\\"job-detail-value\\">No job details or billing information available.</div>";' +
                '  }' +
                '  tooltipContent.innerHTML = html;' +
                '  tooltip.className = "job-details-tooltip visible";' +
                '}' +
                'function hideJobDetailsTooltip() {' +
                '  var tooltip = document.getElementById("jobDetailsTooltip");' +
                '  if (tooltip) tooltip.className = "job-details-tooltip";' +
                '}';
        }

        return {
            onRequest: onRequest
        };
    });
