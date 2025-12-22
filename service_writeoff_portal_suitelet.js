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

define(['N/ui/serverWidget', 'N/query', 'N/log', 'N/url', 'N/record', 'N/runtime'],
    function(serverWidget, query, log, url, record, runtime) {

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
            
            log.debug('handlePost - Request received', {
                action: params.action,
                soId: params.soId,
                allParams: JSON.stringify(params)
            });
            
            try {
                // Check if this is a close action
                if (params.action === 'close' && params.soId) {
                    log.debug('handlePost - Routing to handleCloseSalesOrder', { soId: params.soId });
                    return handleCloseSalesOrder(context);
                }
                
                // Check if this is an auto-bill action
                if (params.action === 'auto-bill' && params.soId) {
                    log.debug('handlePost - Routing to handleAutoBillSalesOrder', { soId: params.soId });
                    return handleAutoBillSalesOrder(context);
                }
                
                // Check if this is a CBSI bill and JE action
                if (params.action === 'cbsi-bill-je' && params.soId) {
                    log.debug('handlePost - Routing to handleCBSIBillAndJE', { soId: params.soId });
                    return handleCBSIBillAndJE(context);
                }
                
                // Check if this is an individual queue action
                if (params.action === 'queue' && params.soId) {
                    log.debug('handlePost - Routing to handleQueueSingleSO', { soId: params.soId });
                    return handleQueueSingleSO(context);
                }
                
                // Check if this is an unqueue action
                if (params.action === 'unqueue' && params.soId) {
                    log.debug('handlePost - Routing to handleUnqueueSalesOrder', { soId: params.soId });
                    return handleUnqueueSalesOrder(context);
                }
                
                // Check if this is an add-note action
                if (params.action === 'add-note' && params.soId) {
                    log.debug('handlePost - Routing to handleAddResearchNote', { soId: params.soId });
                    return handleAddResearchNote(context);
                }
                
                // Get selected SO IDs from checkbox selections
                var selectedSOIds = params.selectedSOIds;
                var bulkAction = params.bulkAction || 'queue'; // Default to queue for backward compatibility
                
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
                
                log.audit('Bulk Action', 'Processing ' + bulkAction + ' for ' + soIdArray.length + ' Sales Orders: ' + soIdArray.join(', '));
                
                // Route to appropriate bulk handler
                if (bulkAction === 'close') {
                    return handleBulkClose(context, soIdArray);
                } else if (bulkAction === 'auto-bill') {
                    return handleBulkAutoBill(context, soIdArray);
                } else if (bulkAction === 'cbsi-bill-je') {
                    return handleBulkCBSI(context, soIdArray);
                }
                
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
         * Extracts a user-friendly error message from a NetSuite error object
         * @param {Error} error - The error object
         * @returns {string} Clean error message
         */
        function getCleanErrorMessage(error) {
            if (!error) return 'Unknown error';
            
            // Try to get the message property first
            var message = error.message || error.toString();
            
            // Common error patterns to simplify
            if (message.indexOf('Address Validation Failed') >= 0) {
                return 'Address validation error: Shipping address contains phone number. Please fix the address on this Sales Order.';
            }
            
            if (message.indexOf('Please enter value(s) for:') >= 0) {
                // Extract the field name
                var match = message.match(/Please enter value\(s\) for: (.+?)(?:\n|$)/);
                if (match) {
                    return 'Missing required field: ' + match[1];
                }
                return message; // Return as-is if we can't parse it
            }
            
            // If the message is very long or contains stack traces, try to get just the first line
            if (message.length > 200 || message.indexOf('\\n') >= 0) {
                var firstLine = message.split('\\n')[0];
                if (firstLine.length > 200) {
                    return firstLine.substring(0, 200) + '...';
                }
                return firstLine;
            }
            
            return message;
        }

        /**
         * Handles closing a Sales Order
         */
        function handleCloseSalesOrder(context) {
            var response = context.response;
            var params = context.request.parameters;
            var soId = params.soId;
            
            log.debug('handleCloseSalesOrder - START', { soId: soId });
            
            try {
                log.debug('handleCloseSalesOrder - Loading SO record', { soId: soId });
                
                // Load and close the sales order
                var soRecord = record.load({
                    type: record.Type.SALES_ORDER,
                    id: soId
                });
                
                log.debug('handleCloseSalesOrder - SO loaded successfully', {
                    soId: soId,
                    currentStatus: soRecord.getValue({ fieldId: 'orderstatus' }),
                    currentStatusText: soRecord.getText({ fieldId: 'orderstatus' })
                });
                
                log.debug('handleCloseSalesOrder - Attempting to close by setting line items', { soId: soId });
                
                // Close all line items instead of setting orderstatus directly
                var lineCount = soRecord.getLineCount({ sublistId: 'item' });
                
                log.debug('handleCloseSalesOrder - Line count', { lineCount: lineCount });
                
                for (var i = 0; i < lineCount; i++) {
                    soRecord.setSublistValue({
                        sublistId: 'item',
                        fieldId: 'isclosed',
                        line: i,
                        value: true
                    });
                }
                
                log.debug('handleCloseSalesOrder - All lines marked as closed, saving...', { soId: soId });
                
                var savedId = soRecord.save();
                
                log.audit('SO Closed', 'Sales Order ' + soId + ' closed successfully. Saved ID: ' + savedId);
                
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ 
                    success: true, 
                    message: 'Sales Order closed successfully.'
                }));
            } catch (e) {
                var errorDetails = {
                    soId: soId,
                    errorString: e.toString(),
                    errorName: e.name || 'Unknown',
                    errorMessage: e.message || 'No message',
                    errorType: typeof e,
                    stack: e.stack || 'No stack trace available'
                };
                
                // Try to extract more error details if available
                try {
                    if (e.cause) errorDetails.cause = e.cause.toString();
                    if (e.id) errorDetails.id = e.id;
                    errorDetails.fullError = JSON.stringify(e);
                } catch (parseErr) {
                    errorDetails.parseError = 'Could not parse full error';
                }
                
                log.error('Close SO Error - DETAILED', errorDetails);
                
                var cleanMessage = getCleanErrorMessage(e);
                log.error('Close SO Error - Simple', 'Failed to close SO ' + soId + ': ' + cleanMessage);
                
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ 
                    success: false, 
                    message: 'Error closing Sales Order: ' + cleanMessage
                }));
            }
        }

        /**
         * Handles auto-billing a Sales Order (transforms to invoice and saves)
         */
        function handleAutoBillSalesOrder(context) {
            var response = context.response;
            var params = context.request.parameters;
            var soId = params.soId;
            
            log.debug('handleAutoBillSalesOrder - START', { soId: soId });
            
            try {
                log.debug('handleAutoBillSalesOrder - Transforming SO to Invoice', { soId: soId });
                
                // Transform SO to Invoice
                var invoiceRecord = record.transform({
                    fromType: record.Type.SALES_ORDER,
                    fromId: soId,
                    toType: record.Type.INVOICE,
                    isDynamic: false
                });
                
                log.debug('handleAutoBillSalesOrder - Invoice transformed, saving...', { soId: soId });
                
                // Save the invoice and capture tranid
                var invoiceId = invoiceRecord.save({
                    enableSourcing: true,
                    ignoreMandatoryFields: false
                });
                
                // Get the invoice tranid
                var invoiceTranid = invoiceRecord.getValue({ fieldId: 'tranid' });
                
                log.audit('SO Auto-Billed', 'Sales Order ' + soId + ' transformed to Invoice ' + invoiceTranid + ' (ID: ' + invoiceId + ')');
                
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ 
                    success: true, 
                    message: 'Invoice created successfully.',
                    invoiceTranid: invoiceTranid,
                    invoiceId: invoiceId
                }));
            } catch (e) {
                var cleanMessage = getCleanErrorMessage(e);
                
                log.error('Auto-Bill Error', {
                    soId: soId,
                    error: e.toString(),
                    errorName: e.name,
                    errorMessage: e.message,
                    cleanMessage: cleanMessage,
                    stack: e.stack || 'No stack trace available'
                });
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ 
                    success: false, 
                    message: 'Error creating invoice: ' + cleanMessage
                }));
            }
        }

        /**
         * Handles CBSI Bill and JE - Creates invoice with entity 335, creates JE, applies JE to invoice
         */
        function handleCBSIBillAndJE(context) {
            var response = context.response;
            var params = context.request.parameters;
            var soId = params.soId;
            
            try {
                log.audit('CBSI Bill and JE Started', 'Processing SO ' + soId);
                
                // STEP 1: Transform SO to Invoice and change entity to 335 (CBSI)
                var invoiceRecord = record.transform({
                    fromType: record.Type.SALES_ORDER,
                    fromId: soId,
                    toType: record.Type.INVOICE,
                    isDynamic: false
                });
                
                // Change entity to 335 (CBSI)
                invoiceRecord.setValue({
                    fieldId: 'entity',
                    value: 335
                });
                
                // Save the invoice
                var invoiceId = invoiceRecord.save({
                    enableSourcing: true,
                    ignoreMandatoryFields: true
                });
                
                // Reload to get accurate values
                invoiceRecord = record.load({
                    type: record.Type.INVOICE,
                    id: invoiceId
                });
                
                var invoiceTranid = invoiceRecord.getValue({ fieldId: 'tranid' });
                var invoiceTotal = Math.abs(parseFloat(invoiceRecord.getValue({ fieldId: 'total' }) || 0));
                
                log.audit('Invoice Created', 'Invoice ' + invoiceTranid + ' (ID: ' + invoiceId + ') with total: $' + invoiceTotal);
                
                // STEP 2: Create Journal Entry
                var jeRecord = record.create({
                    type: record.Type.JOURNAL_ENTRY,
                    isDynamic: false
                });
                
                var jeMemo = 'Automated CBSI Adjustment ' + invoiceTranid;
                jeRecord.setValue({ fieldId: 'memo', value: jeMemo });
                
                // Line 1: Debit Account 470, Dept 13
                jeRecord.setSublistValue({
                    sublistId: 'line',
                    fieldId: 'account',
                    line: 0,
                    value: 470
                });
                jeRecord.setSublistValue({
                    sublistId: 'line',
                    fieldId: 'debit',
                    line: 0,
                    value: invoiceTotal
                });
                jeRecord.setSublistValue({
                    sublistId: 'line',
                    fieldId: 'department',
                    line: 0,
                    value: 13
                });
                jeRecord.setSublistValue({
                    sublistId: 'line',
                    fieldId: 'memo',
                    line: 0,
                    value: jeMemo
                });
                
                // Line 2: Credit Account 119, Entity 335
                jeRecord.setSublistValue({
                    sublistId: 'line',
                    fieldId: 'account',
                    line: 1,
                    value: 119
                });
                jeRecord.setSublistValue({
                    sublistId: 'line',
                    fieldId: 'credit',
                    line: 1,
                    value: invoiceTotal
                });
                jeRecord.setSublistValue({
                    sublistId: 'line',
                    fieldId: 'entity',
                    line: 1,
                    value: 335
                });
                jeRecord.setSublistValue({
                    sublistId: 'line',
                    fieldId: 'memo',
                    line: 1,
                    value: jeMemo
                });
                
                var jeId = jeRecord.save();
                
                // Reload to get tranid
                jeRecord = record.load({
                    type: record.Type.JOURNAL_ENTRY,
                    id: jeId
                });
                var jeTranid = jeRecord.getValue({ fieldId: 'tranid' });
                
                log.audit('JE Created', 'JE ' + jeTranid + ' (ID: ' + jeId + ') with amount: $' + invoiceTotal);
                
                // STEP 3: Apply JE to Invoice using Customer Payment (following reference script pattern)
                var customerPayment = record.transform({
                    fromType: record.Type.INVOICE,
                    fromId: invoiceId,
                    toType: record.Type.CUSTOMER_PAYMENT,
                    isDynamic: false
                });
                
                // Set payment details
                customerPayment.setValue({ fieldId: 'trandate', value: new Date() });
                customerPayment.setValue({ fieldId: 'paymentmethod', value: 15 }); // 15 = ACCT'G
                customerPayment.setValue({ fieldId: 'memo', value: 'CBSI JE Application: ' + jeTranid });
                customerPayment.setValue({ fieldId: 'payment', value: invoiceTotal });
                
                // STEP 4: Clear all auto-selected apply lines
                var applyLineCount = customerPayment.getLineCount({ sublistId: 'apply' });
                
                log.debug('Clearing auto-selected apply lines', { applyLineCount: applyLineCount });
                
                for (var j = 0; j < applyLineCount; j++) {
                    try {
                        var isApplied = customerPayment.getSublistValue({
                            sublistId: 'apply',
                            fieldId: 'apply',
                            line: j
                        });
                        
                        if (isApplied) {
                            customerPayment.setSublistValue({
                                sublistId: 'apply',
                                fieldId: 'apply',
                                line: j,
                                value: false
                            });
                            customerPayment.setSublistValue({
                                sublistId: 'apply',
                                fieldId: 'amount',
                                line: j,
                                value: 0
                            });
                        }
                    } catch (clearError) {
                        log.debug('Could not clear apply line', { line: j, error: clearError.toString() });
                    }
                }
                
                // STEP 5: Find and select the credit transaction (JE)
                var creditLineCount = customerPayment.getLineCount({ sublistId: 'credit' });
                var creditLineUpdated = false;
                var actualCreditAmount = 0;
                
                log.debug('Selecting credit transaction', {
                    jeId: jeId,
                    creditLineCount: creditLineCount,
                    targetAmount: invoiceTotal
                });
                
                for (var c = 0; c < creditLineCount; c++) {
                    var creditDocId = customerPayment.getSublistValue({
                        sublistId: 'credit',
                        fieldId: 'doc',
                        line: c
                    });
                    
                    var creditRefNum = customerPayment.getSublistValue({
                        sublistId: 'credit',
                        fieldId: 'refnum',
                        line: c
                    });
                    
                    if (creditDocId == jeId || creditRefNum == jeId) {
                        try {
                            customerPayment.setSublistValue({
                                sublistId: 'credit',
                                fieldId: 'apply',
                                line: c,
                                value: true
                            });
                            
                            customerPayment.setSublistValue({
                                sublistId: 'credit',
                                fieldId: 'amount',
                                line: c,
                                value: invoiceTotal
                            });
                            
                            actualCreditAmount = customerPayment.getSublistValue({
                                sublistId: 'credit',
                                fieldId: 'amount',
                                line: c
                            });
                            
                            creditLineUpdated = true;
                            
                            log.debug('Selected credit transaction', {
                                line: c,
                                creditDocId: creditDocId,
                                creditRefNum: creditRefNum,
                                targetAmount: invoiceTotal,
                                actualCreditAmount: actualCreditAmount
                            });
                        } catch (creditSetError) {
                            log.error('Error setting credit line', {
                                error: creditSetError.toString(),
                                line: c,
                                creditDocId: creditDocId
                            });
                        }
                        break;
                    }
                }
                
                // STEP 6: Select the invoice for application
                applyLineCount = customerPayment.getLineCount({ sublistId: 'apply' });
                var invoiceLineUpdated = false;
                var actualApplyAmount = 0;
                
                log.debug('Selecting invoice for application', {
                    targetInvoiceId: invoiceId,
                    targetAmount: invoiceTotal,
                    applyLineCount: applyLineCount
                });
                
                for (var j = 0; j < applyLineCount; j++) {
                    var docId = customerPayment.getSublistValue({
                        sublistId: 'apply',
                        fieldId: 'doc',
                        line: j
                    });
                    
                    if (docId == invoiceId) {
                        try {
                            customerPayment.setSublistValue({
                                sublistId: 'apply',
                                fieldId: 'apply',
                                line: j,
                                value: true
                            });
                            
                            customerPayment.setSublistValue({
                                sublistId: 'apply',
                                fieldId: 'amount',
                                line: j,
                                value: invoiceTotal
                            });
                            
                            actualApplyAmount = customerPayment.getSublistValue({
                                sublistId: 'apply',
                                fieldId: 'amount',
                                line: j
                            });
                            
                            invoiceLineUpdated = true;
                            
                            log.debug('Selected invoice for application', {
                                line: j,
                                docId: docId,
                                targetAmount: invoiceTotal,
                                actualApplyAmount: actualApplyAmount
                            });
                        } catch (setError) {
                            log.error('Error selecting invoice line', {
                                error: setError.toString(),
                                line: j,
                                docId: docId
                            });
                        }
                        break;
                    }
                }
                
                // STEP 7: CRITICAL VALIDATION
                var netEffect = actualApplyAmount - actualCreditAmount;
                var amountsMatch = (actualApplyAmount == actualCreditAmount) && (actualApplyAmount == invoiceTotal);
                
                log.debug('FINAL VALIDATION BEFORE SAVE', {
                    expectedAmount: invoiceTotal,
                    actualApplyAmount: actualApplyAmount,
                    actualCreditAmount: actualCreditAmount,
                    netEffect: netEffect,
                    amountsMatch: amountsMatch,
                    invoiceLineUpdated: invoiceLineUpdated,
                    creditLineUpdated: creditLineUpdated
                });
                
                if (!invoiceLineUpdated) {
                    throw new Error('VALIDATION FAILED: Could not select target invoice');
                }
                
                if (!creditLineUpdated) {
                    throw new Error('VALIDATION FAILED: Could not select credit transaction (JE)');
                }
                
                if (!amountsMatch) {
                    throw new Error('VALIDATION FAILED: Amounts do not match. Expected: ' + invoiceTotal + ', Apply: ' + actualApplyAmount + ', Credit: ' + actualCreditAmount);
                }
                
                if (Math.abs(netEffect) > 0.01) {
                    throw new Error('VALIDATION FAILED: Net effect is not zero: ' + netEffect);
                }
                
                // STEP 8: Save the payment to apply the credit
                log.debug('ALL VALIDATIONS PASSED - Saving payment', {
                    paymentAmount: invoiceTotal,
                    applyAmount: actualApplyAmount,
                    creditAmount: actualCreditAmount,
                    netEffect: netEffect
                });
                
                var paymentId = customerPayment.save();
                
                log.debug('Customer payment saved - credit applied successfully', {
                    paymentId: paymentId,
                    netEffect: '$0.00',
                    appliedAmount: actualApplyAmount,
                    creditAmount: actualCreditAmount,
                    invoiceId: invoiceId,
                    jeId: jeId
                });
                
                // STEP 9: DELETE THE PAYMENT RECORD SINCE IT'S NO LONGER NEEDED
                // The credit application has been processed, but we don't need the payment record
                try {
                    log.debug('Deleting temporary payment record', {
                        paymentId: paymentId,
                        reason: 'Credit application complete - payment record not needed'
                    });
                    
                    record.delete({
                        type: record.Type.CUSTOMER_PAYMENT,
                        id: paymentId
                    });
                    
                    log.debug('Temporary payment record deleted successfully', {
                        deletedPaymentId: paymentId,
                        invoiceId: invoiceId,
                        jeId: jeId,
                        result: 'Credit applied without payment record'
                    });
                } catch (deleteError) {
                    log.error('Error deleting temporary payment record', {
                        error: deleteError.toString(),
                        paymentId: paymentId,
                        invoiceId: invoiceId,
                        jeId: jeId,
                        note: 'Credit was applied successfully, but payment record remains'
                    });
                }
                
                log.audit('CBSI Bill and JE Complete', {
                    invoiceTranid: invoiceTranid,
                    invoiceId: invoiceId,
                    jeTranid: jeTranid,
                    jeId: jeId,
                    amount: invoiceTotal
                });
                
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ 
                    success: true, 
                    message: 'CBSI Bill and JE completed successfully.',
                    invoiceTranid: invoiceTranid,
                    jeTranid: jeTranid,
                    amount: invoiceTotal
                }));
                
            } catch (e) {
                var cleanMessage = getCleanErrorMessage(e);
                
                log.error('CBSI Bill and JE Error', 'Failed to process SO ' + soId + ': ' + cleanMessage);
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ 
                    success: false, 
                    message: 'Error in CBSI Bill and JE: ' + cleanMessage
                }));
            }
        }

        /**
         * Handles queueing a single Sales Order for bill & write-off
         */
        function handleQueueSingleSO(context) {
            var response = context.response;
            var params = context.request.parameters;
            var soId = params.soId;
            
            log.debug('handleQueueSingleSO - START', { soId: soId });
            
            try {
                var todayDate = new Date();
                
                log.debug('handleQueueSingleSO - Setting queue date', { soId: soId, date: todayDate });
                
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
                
                log.audit('SO Queued', 'Sales Order ' + soId + ' queued for write-off on ' + todayDate.toISOString());
                
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ 
                    success: true, 
                    message: 'Sales Order queued for Bill & Write-Off processing.',
                    soId: soId
                }));
            } catch (e) {
                var errorDetails = {
                    soId: soId,
                    errorString: e.toString(),
                    errorName: e.name || 'Unknown',
                    errorMessage: e.message || 'No message'
                };
                
                log.error('Queue Single SO Error - DETAILED', errorDetails);
                
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ 
                    success: false, 
                    message: 'Error queueing Sales Order: ' + e.toString() 
                }));
            }
        }

        /**
         * Handles unqueueing a Sales Order (removes the queue date)
         */
        function handleUnqueueSalesOrder(context) {
            var response = context.response;
            var params = context.request.parameters;
            var soId = params.soId;
            
            log.debug('handleUnqueueSalesOrder - START', { soId: soId });
            
            try {
                log.debug('handleUnqueueSalesOrder - Removing queue date', { soId: soId });
                
                record.submitFields({
                    type: record.Type.SALES_ORDER,
                    id: soId,
                    values: {
                        custbody_service_queued_for_write_off: '' // Clear the date
                    },
                    options: {
                        enableSourcing: false,
                        ignoreMandatoryFields: true
                    }
                });
                
                log.audit('SO Unqueued', 'Sales Order ' + soId + ' removed from write-off queue');
                
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ 
                    success: true, 
                    message: 'Sales Order removed from queue.',
                    soId: soId
                }));
            } catch (e) {
                var errorDetails = {
                    soId: soId,
                    errorString: e.toString(),
                    errorName: e.name || 'Unknown',
                    errorMessage: e.message || 'No message'
                };
                
                log.error('Unqueue SO Error - DETAILED', errorDetails);
                
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ 
                    success: false, 
                    message: 'Error unqueueing Sales Order: ' + e.toString() 
                }));
            }
        }

        /**
         * Handles adding a research note to a Sales Order
         */
        function handleAddResearchNote(context) {
            var response = context.response;
            var params = context.request.parameters;
            var soId = params.soId;
            var note = params.note || '';
            var followUpDateParam = params.followUpDate || '';
            
            // Convert date from YYYY-MM-DD (HTML date input format) to M/D/YYYY (NetSuite format)
            var followUpDate = '';
            if (followUpDateParam) {
                try {
                    var dateParts = followUpDateParam.split('-');
                    if (dateParts.length === 3) {
                        var year = dateParts[0];
                        var month = parseInt(dateParts[1], 10); // Remove leading zero
                        var day = parseInt(dateParts[2], 10); // Remove leading zero
                        followUpDate = month + '/' + day + '/' + year;
                    }
                } catch (dateErr) {
                    log.error('Date Conversion Error', { original: followUpDateParam, error: dateErr.toString() });
                    followUpDate = followUpDateParam; // Use original if conversion fails
                }
            }
            
            log.debug('handleAddResearchNote - START', { soId: soId, noteLength: note.length, followUpDateParam: followUpDateParam, followUpDate: followUpDate });
            
            try {
                record.submitFields({
                    type: record.Type.SALES_ORDER,
                    id: soId,
                    values: {
                        custbody_service_research_notes: note,
                        custbody_service_research_followupdate: followUpDate
                    },
                    options: {
                        enableSourcing: false,
                        ignoreMandatoryFields: true
                    }
                });
                
                log.audit('Research Note Added', 'Sales Order ' + soId + ' - Note: ' + (note.substring(0, 50) || '(cleared)') + ' - Follow Up: ' + (followUpDate || '(cleared)'));
                
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ 
                    success: true, 
                    message: note ? 'Research note saved.' : 'Research note cleared.',
                    soId: soId,
                    followUpDate: followUpDateParam
                }));
            } catch (e) {
                var errorDetails = {
                    soId: soId,
                    errorString: e.toString(),
                    errorName: e.name || 'Unknown',
                    errorMessage: e.message || 'No message'
                };
                
                log.error('Add Research Note Error - DETAILED', errorDetails);
                
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ 
                    success: false, 
                    message: 'Error saving research note: ' + e.toString() 
                }));
            }
        }

        /**
         * Handles bulk close for multiple Sales Orders
         */
        function handleBulkClose(context, soIdArray) {
            var response = context.response;
            var processedIds = [];
            var failedIds = [];
            var failureDetails = {}; // Track error messages per SO
            var governanceStopped = false;
            var GOVERNANCE_THRESHOLD = 50;
            
            log.debug('handleBulkClose - START', { count: soIdArray.length });
            
            for (var i = 0; i < soIdArray.length; i++) {
                var soId = soIdArray[i];
                
                // Check governance
                var remainingUsage = runtime.getCurrentScript().getRemainingUsage();
                log.debug('Governance Check', { iteration: i, soId: soId, remaining: remainingUsage });
                
                if (remainingUsage < GOVERNANCE_THRESHOLD) {
                    log.audit('Governance Limit Approaching', {
                        processed: processedIds.length,
                        remaining: soIdArray.length - i,
                        governanceRemaining: remainingUsage
                    });
                    governanceStopped = true;
                    break;
                }
                
                try {
                    log.debug('Bulk Close - Loading SO', { soId: soId });
                    
                    var soRecord = record.load({
                        type: record.Type.SALES_ORDER,
                        id: soId
                    });
                    
                    log.debug('Bulk Close - SO Loaded', { 
                        soId: soId, 
                        status: soRecord.getValue({ fieldId: 'status' }),
                        statusText: soRecord.getText({ fieldId: 'status' })
                    });
                    
                    var lineCount = soRecord.getLineCount({ sublistId: 'item' });
                    log.debug('Bulk Close - Processing lines', { soId: soId, lineCount: lineCount });
                    
                    for (var j = 0; j < lineCount; j++) {
                        var lineIsClosed = soRecord.getSublistValue({
                            sublistId: 'item',
                            fieldId: 'isclosed',
                            line: j
                        });
                        
                        log.debug('Bulk Close - Line status', { 
                            soId: soId, 
                            line: j, 
                            currentlyClosed: lineIsClosed 
                        });
                        
                        soRecord.setSublistValue({
                            sublistId: 'item',
                            fieldId: 'isclosed',
                            line: j,
                            value: true
                        });
                    }
                    
                    log.debug('Bulk Close - About to save SO', { soId: soId, lineCount: lineCount });
                    
                    var savedId = soRecord.save({
                        enableSourcing: false,
                        ignoreMandatoryFields: true
                    });
                    
                    log.debug('Bulk Close - IMMEDIATELY after save() returned', { soId: soId, savedId: savedId, typeOfSavedId: typeof savedId });
                    
                    processedIds.push(soId);
                    log.audit('Bulk Close - SO Closed Successfully', { soId: soId, savedId: savedId, linesClosed: lineCount });
                } catch (e) {
                    log.debug('Bulk Close - CATCH BLOCK ENTERED', { soId: soId, errorType: typeof e });
                    
                    failedIds.push(soId);
                    
                    // Extract user-friendly error message
                    var userMessage = getCleanErrorMessage(e);
                    failureDetails[soId] = userMessage;
                    
                    var errorDetails = {
                        soId: soId,
                        errorString: e.toString(),
                        errorName: e.name || 'Unknown',
                        errorMessage: e.message || 'No message',
                        errorStack: e.stack || 'No stack trace',
                        errorType: e.type || 'Unknown type'
                    };
                    
                    if (e.cause) {
                        errorDetails.errorCause = e.cause;
                    }
                    
                    log.debug('Bulk Close Error - DETAILED', errorDetails);
                }
            }
            
            log.debug('Bulk Close - Loop complete', { 
                processedCount: processedIds.length, 
                failedCount: failedIds.length,
                processedIds: processedIds.join(','),
                failedIds: failedIds.join(','),
                governanceStopped: governanceStopped
            });
            
            var message = processedIds.length + ' Sales Order(s) closed.';
            if (failedIds.length > 0) {
                message += ' Failed: ' + failedIds.length;
                // Add details about failures
                var errorSummary = [];
                for (var i = 0; i < failedIds.length; i++) {
                    var fid = failedIds[i];
                    errorSummary.push('SO #' + fid + ': ' + (failureDetails[fid] || 'Unknown error'));
                }
                message += '\\n\\nFailure details:\\n' + errorSummary.join('\\n');
            }
            if (governanceStopped) {
                message += '\\n\\nGOVERNANCE LIMIT: Processed ' + processedIds.length + ' of ' + soIdArray.length + '. Remaining items still selected - click again to continue.';
            }
            
            response.setHeader({ name: 'Content-Type', value: 'application/json' });
            response.write(JSON.stringify({
                success: true,
                message: message,
                processedIds: processedIds,
                failedIds: failedIds,
                failureDetails: failureDetails,
                governanceStopped: governanceStopped,
                count: processedIds.length
            }));
        }

        /**
         * Handles bulk auto-bill for multiple Sales Orders
         */
        function handleBulkAutoBill(context, soIdArray) {
            var response = context.response;
            var processedIds = [];
            var failedIds = [];
            var invoiceDetails = [];
            var governanceStopped = false;
            var GOVERNANCE_THRESHOLD = 50;
            
            log.debug('handleBulkAutoBill - START', { count: soIdArray.length });
            
            for (var i = 0; i < soIdArray.length; i++) {
                var soId = soIdArray[i];
                
                // Check governance
                var remainingUsage = runtime.getCurrentScript().getRemainingUsage();
                log.debug('Governance Check', { iteration: i, soId: soId, remaining: remainingUsage });
                
                if (remainingUsage < GOVERNANCE_THRESHOLD) {
                    log.audit('Governance Limit Approaching', {
                        processed: processedIds.length,
                        remaining: soIdArray.length - i,
                        governanceRemaining: remainingUsage
                    });
                    governanceStopped = true;
                    break;
                }
                
                try {
                    log.debug('Bulk Auto-Bill - Transforming SO to Invoice', { soId: soId });
                    
                    var invoiceRecord = record.transform({
                        fromType: record.Type.SALES_ORDER,
                        fromId: soId,
                        toType: record.Type.INVOICE,
                        isDynamic: false
                    });
                    
                    log.debug('Bulk Auto-Bill - Invoice transformed, saving', { soId: soId });
                    var invoiceId = invoiceRecord.save();
                    var invoiceTranid = invoiceRecord.getValue({ fieldId: 'tranid' });
                    
                    processedIds.push(soId);
                    invoiceDetails.push({ soId: soId, invoiceTranid: invoiceTranid, invoiceId: invoiceId });
                    log.audit('Bulk Auto-Bill - Invoice Created Successfully', { soId: soId, invoiceTranid: invoiceTranid, invoiceId: invoiceId });
                } catch (e) {
                    failedIds.push(soId);
                    
                    var errorDetails = {
                        soId: soId,
                        errorString: e.toString(),
                        errorName: e.name || 'Unknown',
                        errorMessage: e.message || 'No message',
                        errorStack: e.stack || 'No stack trace',
                        errorType: e.type || 'Unknown type'
                    };
                    
                    if (e.cause) {
                        errorDetails.errorCause = e.cause;
                    }
                    
                    log.error('Bulk Auto-Bill Error - DETAILED', errorDetails);
                }
            }
            
            var message = processedIds.length + ' Invoice(s) created.';
            if (failedIds.length > 0) {
                message += ' Failed: ' + failedIds.length;
            }
            if (governanceStopped) {
                message += ' GOVERNANCE LIMIT: Processed ' + processedIds.length + ' of ' + soIdArray.length + '. Remaining items still selected - click again to continue.';
            }
            
            response.setHeader({ name: 'Content-Type', value: 'application/json' });
            response.write(JSON.stringify({
                success: true,
                message: message,
                processedIds: processedIds,
                failedIds: failedIds,
                invoiceDetails: invoiceDetails,
                governanceStopped: governanceStopped,
                count: processedIds.length
            }));
        }

        /**
         * Handles bulk CBSI Bill and JE for multiple Sales Orders
         */
        function handleBulkCBSI(context, soIdArray) {
            var response = context.response;
            var processedIds = [];
            var failedIds = [];
            var cbsiDetails = [];
            var governanceStopped = false;
            var GOVERNANCE_THRESHOLD = 150; // Higher threshold for CBSI due to complexity
            
            log.debug('handleBulkCBSI - START', { count: soIdArray.length });
            
            for (var i = 0; i < soIdArray.length; i++) {
                var soId = soIdArray[i];
                
                // Check governance
                var remainingUsage = runtime.getCurrentScript().getRemainingUsage();
                log.debug('Governance Check', { iteration: i, soId: soId, remaining: remainingUsage });
                
                if (remainingUsage < GOVERNANCE_THRESHOLD) {
                    log.audit('Governance Limit Approaching', {
                        processed: processedIds.length,
                        remaining: soIdArray.length - i,
                        governanceRemaining: remainingUsage
                    });
                    governanceStopped = true;
                    break;
                }
                
                try {
                    // This is a simplified version - full CBSI logic from handleCBSIBillAndJE
                    var invoiceRecord = record.transform({
                        fromType: record.Type.SALES_ORDER,
                        fromId: soId,
                        toType: record.Type.INVOICE,
                        isDynamic: false
                    });
                    
                    invoiceRecord.setValue({ fieldId: 'entity', value: 335 });
                    var invoiceId = invoiceRecord.save({ ignoreMandatoryFields: true });
                    
                    invoiceRecord = record.load({ type: record.Type.INVOICE, id: invoiceId });
                    var invoiceTranid = invoiceRecord.getValue({ fieldId: 'tranid' });
                    var invoiceTotal = Math.abs(parseFloat(invoiceRecord.getValue({ fieldId: 'total' }) || 0));
                    
                    // Create JE (simplified - keeping essential parts)
                    var jeRecord = record.create({ type: record.Type.JOURNAL_ENTRY, isDynamic: false });
                    var jeMemo = 'Automated CBSI Adjustment ' + invoiceTranid;
                    jeRecord.setValue({ fieldId: 'memo', value: jeMemo });
                    
                    jeRecord.setSublistValue({ sublistId: 'line', fieldId: 'account', line: 0, value: 470 });
                    jeRecord.setSublistValue({ sublistId: 'line', fieldId: 'debit', line: 0, value: invoiceTotal });
                    jeRecord.setSublistValue({ sublistId: 'line', fieldId: 'department', line: 0, value: 13 });
                    jeRecord.setSublistValue({ sublistId: 'line', fieldId: 'memo', line: 0, value: jeMemo });
                    
                    jeRecord.setSublistValue({ sublistId: 'line', fieldId: 'account', line: 1, value: 119 });
                    jeRecord.setSublistValue({ sublistId: 'line', fieldId: 'credit', line: 1, value: invoiceTotal });
                    jeRecord.setSublistValue({ sublistId: 'line', fieldId: 'entity', line: 1, value: 335 });
                    jeRecord.setSublistValue({ sublistId: 'line', fieldId: 'memo', line: 1, value: jeMemo });
                    
                    var jeId = jeRecord.save();
                    jeRecord = record.load({ type: record.Type.JOURNAL_ENTRY, id: jeId });
                    var jeTranid = jeRecord.getValue({ fieldId: 'tranid' });
                    
                    // Apply JE to Invoice (simplified - essential parts only)
                    var customerPayment = record.transform({
                        fromType: record.Type.INVOICE,
                        fromId: invoiceId,
                        toType: record.Type.CUSTOMER_PAYMENT,
                        isDynamic: false
                    });
                    
                    customerPayment.setValue({ fieldId: 'trandate', value: new Date() });
                    customerPayment.setValue({ fieldId: 'paymentmethod', value: 15 });
                    customerPayment.setValue({ fieldId: 'memo', value: 'CBSI JE Application: ' + jeTranid });
                    customerPayment.setValue({ fieldId: 'payment', value: invoiceTotal });
                    
                    // Clear auto-selected applies
                    var applyLineCount = customerPayment.getLineCount({ sublistId: 'apply' });
                    for (var j = 0; j < applyLineCount; j++) {
                        try {
                            customerPayment.setSublistValue({ sublistId: 'apply', fieldId: 'apply', line: j, value: false });
                            customerPayment.setSublistValue({ sublistId: 'apply', fieldId: 'amount', line: j, value: 0 });
                        } catch (e) { /* ignore */ }
                    }
                    
                    // Select credit (JE)
                    var creditLineCount = customerPayment.getLineCount({ sublistId: 'credit' });
                    for (var c = 0; c < creditLineCount; c++) {
                        var creditDocId = customerPayment.getSublistValue({ sublistId: 'credit', fieldId: 'doc', line: c });
                        if (creditDocId == jeId) {
                            customerPayment.setSublistValue({ sublistId: 'credit', fieldId: 'apply', line: c, value: true });
                            customerPayment.setSublistValue({ sublistId: 'credit', fieldId: 'amount', line: c, value: invoiceTotal });
                            break;
                        }
                    }
                    
                    // Select invoice
                    applyLineCount = customerPayment.getLineCount({ sublistId: 'apply' });
                    for (var j = 0; j < applyLineCount; j++) {
                        var docId = customerPayment.getSublistValue({ sublistId: 'apply', fieldId: 'doc', line: j });
                        if (docId == invoiceId) {
                            customerPayment.setSublistValue({ sublistId: 'apply', fieldId: 'apply', line: j, value: true });
                            customerPayment.setSublistValue({ sublistId: 'apply', fieldId: 'amount', line: j, value: invoiceTotal });
                            break;
                        }
                    }
                    
                    var paymentId = customerPayment.save();
                    
                    // Delete temp payment
                    try {
                        record.delete({ type: record.Type.CUSTOMER_PAYMENT, id: paymentId });
                    } catch (e) { /* ignore deletion errors */ }
                    
                    processedIds.push(soId);
                    cbsiDetails.push({ soId: soId, invoiceTranid: invoiceTranid, jeTranid: jeTranid, amount: invoiceTotal });
                    log.audit('Bulk CBSI - Complete Successfully', { soId: soId, invoiceTranid: invoiceTranid, jeTranid: jeTranid, amount: invoiceTotal });
                } catch (e) {
                    failedIds.push(soId);
                    
                    var errorDetails = {
                        soId: soId,
                        errorString: e.toString(),
                        errorName: e.name || 'Unknown',
                        errorMessage: e.message || 'No message',
                        errorStack: e.stack || 'No stack trace',
                        errorType: e.type || 'Unknown type'
                    };
                    
                    if (e.cause) {
                        errorDetails.errorCause = e.cause;
                    }
                    
                    log.error('Bulk CBSI Error - DETAILED', errorDetails);
                }
            }
            
            var message = processedIds.length + ' CBSI transactions completed.';
            if (failedIds.length > 0) {
                message += ' Failed: ' + failedIds.length;
            }
            if (governanceStopped) {
                message += ' GOVERNANCE LIMIT: Processed ' + processedIds.length + ' of ' + soIdArray.length + '. Remaining items still selected - click again to continue.';
            }
            
            response.setHeader({ name: 'Content-Type', value: 'application/json' });
            response.write(JSON.stringify({
                success: true,
                message: message,
                processedIds: processedIds,
                failedIds: failedIds,
                cbsiDetails: cbsiDetails,
                governanceStopped: governanceStopped,
                count: processedIds.length
            }));
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
            
            log.debug('handleLoadData - START', 'Loading unbilled SO data via AJAX');
            log.audit('Service Write-Off Portal', 'Loading unbilled SO data via AJAX');

            try {
                log.debug('handleLoadData - Calling runMainQuery');
                
                // Run the main query - returns one row per SO with aggregated unbilled data
                var salesOrders = runMainQuery();
                
                log.debug('handleLoadData - Query complete', { resultCount: salesOrders.length });
                log.audit('Query Complete', 'Found ' + salesOrders.length + ' Sales Orders with unbilled items');

                // Calculate summary stats for ALL data and QUEUED data
                var totalSOs = salesOrders.length;
                var totalUnbilledLineCount = 0;
                var totalUnbilledAmount = 0;
                var queuedSOs = 0;
                var queuedUnbilledLineCount = 0;
                var queuedUnbilledAmount = 0;
                
                for (var i = 0; i < salesOrders.length; i++) {
                    var so = salesOrders[i];
                    var unbilledLines = parseInt(so.unbilled_line_count || 0);
                    var unbilledAmt = parseFloat(so.total_unbilled_amount || 0) * -1;
                    
                    totalUnbilledLineCount += unbilledLines;
                    totalUnbilledAmount += unbilledAmt;
                    
                    // Track queued totals
                    if (so.queued_date) {
                        queuedSOs++;
                        queuedUnbilledLineCount += unbilledLines;
                        queuedUnbilledAmount += unbilledAmt;
                    }
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
                    queuedTotal: queuedSOs,
                    queuedTotalLines: queuedUnbilledLineCount,
                    queuedTotalAmount: queuedUnbilledAmount,
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
            log.debug('runMainQuery - START', 'Building and executing SuiteQL query');
            
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
                "so.custbody_service_research_notes AS research_notes, " +
                "so.custbody_service_research_followupdate AS follow_up_date, " +
                "MAX(BUILTIN.DF(so.custbody_bas_fa_parts_status)) AS parts_status, " +
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
                "    INNER JOIN transaction inv ON inv_line.transaction = inv.id " +
                "    WHERE inv_line.createdfrom = so.id " +
                "    AND inv_line.item = so_line.item " +
                "    AND inv.type = 'CustInvc' " +
                "    AND inv_line.taxline = 'F' " +
                "    AND inv_line.mainline = 'F' " +
                ") " +
                "GROUP BY so.id, so.tranid, so.trandate, so.entity, so.status, so.custbody_f4n_job_id, so.custbody_service_queued_for_write_off, so.custbody24, so.shipdate, so.custbody_bas_estimated_ship_date, so.custbody_f4n_details, so.custbody21, so.custbody_f4n_job_state, so.custbody_f4n_scheduled, so.custbody_f4n_started, so.custbody_f4n_completed, so.custbody_service_research_notes, so.custbody_service_research_followupdate " +
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
                '<div id="researchNoteModal" class="modal-overlay" style="display:none;">' +
                '<div class="modal-content">' +
                '<div class="modal-header">📝 Research Note</div>' +
                '<div class="modal-body">' +
                '<label class="modal-label">Enter research note for this Sales Order:</label>' +
                '<textarea id="researchNoteInput" class="modal-textarea" rows="6" placeholder="Enter notes here..."></textarea>' +
                '<label class="modal-label" style="margin-top: 15px;">Follow Up Date:</label>' +
                '<input type="date" id="followUpDateInput" class="modal-input" />' +
                '</div>' +
                '<div class="modal-footer">' +
                '<button type="button" class="modal-btn modal-btn-cancel" onclick="closeResearchNoteModal()">Cancel</button>' +
                '<button type="button" class="modal-btn modal-btn-save" onclick="saveResearchNote()">Save Note</button>' +
                '</div>' +
                '</div>' +
                '</div>' +
                '</div>';

            return html;
        }

        /**
         * Builds summary statistics section with three horizontal summaries (all, queued, selected)
         * @param {Array} data - Sales Order data
         * @returns {string} Summary HTML
         */
        function buildSummarySection(data) {
            return '<div class="summary-container">' +
                '<div class="summary-section">' +
                '<h3 class="summary-section-header">📊 All Unbilled</h3>' +
                '<div class="summary-card">' +
                '<div class="summary-value" id="summaryTotal">0</div>' +
                '<div class="summary-label">Total SOs</div>' +
                '</div>' +
                '<div class="summary-card card-pending">' +
                '<div class="summary-value" id="summaryTotalLines">0</div>' +
                '<div class="summary-label">Unbilled Lines</div>' +
                '</div>' +
                '<div class="summary-card card-open">' +
                '<div class="summary-value" id="summaryTotalAmount">$0.00</div>' +
                '<div class="summary-label">Unbilled Amount</div>' +
                '</div>' +
                '</div>' +
                '<div class="summary-section">' +
                '<h3 class="summary-section-header">⏳ Queued for Write-Off</h3>' +
                '<div class="summary-card card-queued">' +
                '<div class="summary-value" id="queuedTotal">0</div>' +
                '<div class="summary-label">Queued SOs</div>' +
                '</div>' +
                '<div class="summary-card card-queued-lines">' +
                '<div class="summary-value" id="queuedTotalLines">0</div>' +
                '<div class="summary-label">Queued Lines</div>' +
                '</div>' +
                '<div class="summary-card card-queued-amount">' +
                '<div class="summary-value" id="queuedTotalAmount">$0.00</div>' +
                '<div class="summary-label">Queued Amount</div>' +
                '</div>' +
                '</div>' +
                '<div class="summary-section">' +
                '<h3 class="summary-section-header">✅ Selected</h3>' +
                '<div class="summary-card card-selected">' +
                '<div class="summary-value" id="selectedCount">0</div>' +
                '<div class="summary-label">Selected SOs</div>' +
                '</div>' +
                '<div class="summary-card card-selected-lines">' +
                '<div class="summary-value" id="selectedLines">0</div>' +
                '<div class="summary-label">Selected Lines</div>' +
                '</div>' +
                '<div class="summary-card card-selected-amount">' +
                '<div class="summary-value" id="selectedAmount">$0.00</div>' +
                '<div class="summary-label">Selected Amount</div>' +
                '</div>' +
                '</div>' +
                '<div class="filter-section">' +
                '<h3 class="filter-section-header">🔍 Filters</h3>' +
                '<div class="filter-cards-row">' +
                '<div class="filter-card">' +
                '<div class="filter-subsection-header">Ship Dates</div>' +
                '<label class="filter-item">' +
                '<input type="checkbox" id="filterNoShipDate" checked onchange="applyFilters()">' +
                '<span class="filter-label">No Ship Date</span>' +
                '</label>' +
                '<label class="filter-item">' +
                '<input type="checkbox" id="filterOldShipDates" checked onchange="applyFilters()">' +
                '<span class="filter-label">Old Ship Dates</span>' +
                '</label>' +
                '<label class="filter-item">' +
                '<input type="checkbox" id="filterFutureShipDates" onchange="applyFilters()">' +
                '<span class="filter-label">Future Ship Dates</span>' +
                '</label>' +
                '<div class="filter-subsection-header" style="margin-top: 12px; border-top: 1px solid #cbd5e1; padding-top: 12px;">Research Notes</div>' +
                '<label class="filter-item">' +
                '<input type="checkbox" id="filterHasResearchNotes" checked onchange="applyFilters()">' +
                '<span class="filter-label">Has Research Notes</span>' +
                '</label>' +
                '<label class="filter-item">' +
                '<input type="checkbox" id="filterNoResearchNotes" checked onchange="applyFilters()">' +
                '<span class="filter-label">No Research Notes</span>' +
                '</label>' +
                '</div>' +
                '<div class="filter-card">' +
                '<div class="filter-subsection-header">Job State</div>' +
                '<label class="filter-item">' +
                '<input type="checkbox" id="filterJobScheduled" checked onchange="applyFilters()">' +
                '<span class="filter-label">Scheduled</span>' +
                '</label>' +
                '<label class="filter-item">' +
                '<input type="checkbox" id="filterJobActive" checked onchange="applyFilters()">' +
                '<span class="filter-label">Active</span>' +
                '</label>' +
                '<label class="filter-item">' +
                '<input type="checkbox" id="filterJobPaused" checked onchange="applyFilters()">' +
                '<span class="filter-label">Paused</span>' +
                '</label>' +
                '<label class="filter-item">' +
                '<input type="checkbox" id="filterJobCompleted" checked onchange="applyFilters()">' +
                '<span class="filter-label">Completed</span>' +
                '</label>' +
                '</div>' +
                '<div class="filter-card">' +
                '<div class="filter-subsection-header">Warranty Type</div>' +
                '<label class="filter-item">' +
                '<input type="checkbox" id="filterWarrantyNone" checked onchange="applyFilters()">' +
                '<span class="filter-label">No Warranty Type</span>' +
                '</label>' +
                '<label class="filter-item">' +
                '<input type="checkbox" id="filterWarrantyCBSI" checked onchange="applyFilters()">' +
                '<span class="filter-label">CBSI</span>' +
                '</label>' +
                '<label class="filter-item">' +
                '<input type="checkbox" id="filterWarrantyCOD" checked onchange="applyFilters()">' +
                '<span class="filter-label">Cash on Delivery</span>' +
                '</label>' +
                '<label class="filter-item">' +
                '<input type="checkbox" id="filterWarrantyExtended" checked onchange="applyFilters()">' +
                '<span class="filter-label">Extended Warranty</span>' +
                '</label>' +
                '<label class="filter-item">' +
                '<input type="checkbox" id="filterWarrantyMfg" checked onchange="applyFilters()">' +
                '<span class="filter-label">Mfg Warranty Term</span>' +
                '</label>' +
                '<label class="filter-item">' +
                '<input type="checkbox" id="filterWarrantyShop" checked onchange="applyFilters()">' +
                '<span class="filter-label">Shop Ticket</span>' +
                '</label>' +
                '</div>' +
                '</div>' +
                '<div class="filter-showing-count" id="filterShipDateCount">Showing: 0 / 0</div>' +
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
                '<div class="bulk-action-container">' +
                '<select id="bulkActionSelect" class="bulk-action-select" disabled>' +
                '<option value="">▼ Select Bulk Actions</option>' +
                '<option value="queue">Queue for Bill & Write-Off</option>' +
                '<option value="close">Close (Cancel)</option>' +
                '<option value="auto-bill">Auto-Bill (Invoice)</option>' +
                '<option value="cbsi-bill-je">CBSI (Bill and JE)</option>' +
                '</select>' +
                '<button type="button" id="bulkActionBtn" class="action-btn-large" onclick="executeBulkAction()" disabled>Apply</button>' +
                '</div>' +
                '</div>' +
                '<div class="table-wrapper">' +
                '<table id="dataTable">' +
                '<thead>' +
                '<tr>' +
                '<th class="th-checkbox"><input type="checkbox" id="selectAllCheckbox" onchange="toggleSelectAll(this)"></th>' +
                '<th class="th-actions">Actions</th>' +
                '<th class="th-slate" onclick="sortTable(2)">Sales Order<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate th-follow-up" onclick="sortTable(3)">Follow Up<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(4)">Queued<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(5)">Job ID<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(6)">Warranty Type<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(7)">EPIC Auth<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(8)">Customer<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(9)">SO Date<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(10)">Ship Date<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(11)">Est Ship Date<span class="sort-arrow">↕</span></th>' +
                '<th class="th-slate" onclick="sortTable(12)">SO Status<span class="sort-arrow">↕</span></th>' +
                '<th class="th-teal" onclick="sortTable(13)">Unbilled<br>Line Count<span class="sort-arrow">↕</span></th>' +
                '<th class="th-teal" onclick="sortTable(14)">Total Unbilled<br>Amount<span class="sort-arrow">↕</span></th>' +
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
            var partsStatus = record.parts_status || '';
            var scheduledDate = record.scheduled_date || '';
            var jobStarted = record.job_started || '';
            var jobCompleted = record.job_completed || '';
            var queuedDate = record.queued_date || '';
            
            var shipDateRaw = record.ship_date || '';
            var warrantyType = record.warranty_type || '';
            var researchNotes = record.research_notes || '';
            var followUpDate = record.follow_up_date || '';
            var followUpDateForInput = toInputDateFormat(followUpDate); // Convert to YYYY-MM-DD for date input
            var html = '<tr class="line-items-row" data-so-id="' + soId + '" data-unbilled-lines="' + unbilledLines + '" data-unbilled-amount="' + unbilledAmount + '" data-unbilled-detail="' + unbilledDetail + '" data-job-details="' + escapeHtml(jobDetails) + '" data-billing-completed-by="' + escapeHtml(billingCompletedBy) + '" data-job-state="' + escapeHtml(jobState) + '" data-parts-status="' + escapeHtml(partsStatus) + '" data-scheduled-date="' + scheduledDate + '" data-job-started="' + jobStarted + '" data-job-completed="' + jobCompleted + '" data-ship-date="' + shipDateRaw + '" data-warranty-type="' + escapeHtml(warrantyType) + '" data-research-notes="' + escapeHtml(researchNotes) + '" data-follow-up-date="' + followUpDateForInput + '" onmouseenter="showLineItemsTooltip(this); showJobDetailsTooltip(this);" onmouseleave="hideLineItemsTooltip(); hideJobDetailsTooltip();">';
            
            // Checkbox column
            html += '<td class="col-checkbox"><input type="checkbox" class="so-checkbox" value="' + soId + '" onchange="updateSelectedSummary()"></td>';
            
            // Actions dropdown
            html += '<td class="col-actions"><select class="actions-dropdown" onchange="handleAction(this, ' + soId + ')" onclick="event.stopPropagation();"><option value="">▼</option><option value="queue">Queue for Bill & Write-Off</option><option value="close">Close (Cancel)</option><option value="bill">Manual Bill (Invoice)</option><option value="auto-bill">Auto-Bill (Invoice)</option><option value="cbsi-bill-je">CBSI (Bill and JE)</option><option value="add-note">Add Research Note</option></select></td>';
            
            // Sales Order columns (slate group)
            var noteIcon = researchNotes ? ' <span class="research-note-icon" id="note-icon-' + soId + '" title="Has research notes">ℹ️</span>' : '<span class="research-note-icon" id="note-icon-' + soId + '" style="display:none;" title="Has research notes">ℹ️</span>';
            html += '<td class="col-slate">' + buildTransactionLink(record.so_id, record.so_number, 'salesord') + noteIcon + '</td>';
            html += '<td class="col-slate follow-up-cell" id="follow-up-cell-' + soId + '">' + formatDate(followUpDateForInput) + '</td>';
            html += '<td class="col-slate queued-cell" id="queued-cell-' + soId + '">';
            if (queuedDate) {
                html += '<span class="queued-checkmark">✓</span><span class="unqueue-x" onclick="handleUnqueue(' + soId + ', event)" title="Remove from queue">✕</span>';
            }
            html += '</td>';
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
                // If date is in YYYY-MM-DD format, parse it as local date to avoid timezone issues
                if (typeof dateValue === 'string' && dateValue.match(/^\d{4}-\d{2}-\d{2}$/)) {
                    var parts = dateValue.split('-');
                    var year = parseInt(parts[0], 10);
                    var month = parseInt(parts[1], 10) - 1; // Month is 0-indexed
                    var day = parseInt(parts[2], 10);
                    var d = new Date(year, month, day);
                    return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
                }
                var d = new Date(dateValue);
                return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
            } catch (e) {
                return dateValue;
            }
        }

        /**
         * Converts a date to YYYY-MM-DD format for HTML date inputs
         * @param {string} dateValue
         * @returns {string} Date in YYYY-MM-DD format
         */
        function toInputDateFormat(dateValue) {
            if (!dateValue) return '';
            try {
                var d;
                // If already in YYYY-MM-DD format, return as-is
                if (typeof dateValue === 'string' && dateValue.match(/^\d{4}-\d{2}-\d{2}$/)) {
                    return dateValue;
                }
                d = new Date(dateValue);
                if (isNaN(d.getTime())) return '';
                var year = d.getFullYear();
                var month = String(d.getMonth() + 1).padStart(2, '0');
                var day = String(d.getDate()).padStart(2, '0');
                return year + '-' + month + '-' + day;
            } catch (e) {
                return '';
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
                /* Summary container - four columns (3 summary sections + filters) */
                '.summary-container { display: grid; grid-template-columns: minmax(100px, 1fr) minmax(100px, 1fr) minmax(100px, 1fr) 500px; gap: 10px; margin: 20px 0 30px 0; }' +
                '.summary-section { display: flex; flex-direction: column; gap: 10px; }' +
                '.summary-section-header { color: #013220; font-size: 16px; font-weight: 600; margin: 0 0 10px 0; padding-bottom: 6px; border-bottom: 2px solid #e2e8f0; text-align: center; }' +
                '.summary-card { background: #E6EEEA; padding: 12px; border-radius: 6px; text-align: center; border-left: 4px solid #013220; }' +
                '.summary-card.card-pending { border-left-color: #014421; background: #E8F2EC; }' +
                '.summary-card.card-open { border-left-color: #355E3B; background: #EBF0EB; }' +
                '.summary-card.card-queued { border-left-color: #B8860B; background: #FFF8DC; }' +
                '.summary-card.card-queued-lines { border-left-color: #DAA520; background: #FFFACD; }' +
                '.summary-card.card-queued-amount { border-left-color: #CD853F; background: #FAEBD7; }' +
                '.summary-card.card-selected { border-left-color: #8A9A5B; background: #F4F7F0; }' +
                '.summary-card.card-selected-lines { border-left-color: #6B7F3F; background: #F2F5ED; }' +
                '.summary-card.card-selected-amount { border-left-color: #556B2F; background: #F0F3EC; }' +
                '.summary-value { font-size: 24px; font-weight: bold; color: #013220; margin-bottom: 4px; }' +
                '.summary-card.card-pending .summary-value { color: #014421; }' +
                '.summary-card.card-open .summary-value { color: #355E3B; }' +
                '.summary-card.card-queued .summary-value { color: #B8860B; }' +
                '.summary-card.card-queued-lines .summary-value { color: #DAA520; }' +
                '.summary-card.card-queued-amount .summary-value { color: #CD853F; }' +
                '.summary-card.card-selected .summary-value { color: #8A9A5B; }' +
                '.summary-card.card-selected-lines .summary-value { color: #6B7F3F; }' +
                '.summary-card.card-selected-amount .summary-value { color: #556B2F; }' +
                '.summary-label { font-size: 12px; color: #2d3a33; font-weight: 600; }' +
                '.summary-sublabel { font-size: 11px; color: #6b7c72; margin-top: 4px; font-style: italic; }' +
                /* Filter section */
                '.filter-section { display: flex; flex-direction: column; gap: 10px; }' +
                '.filter-section-header { color: #013220; font-size: 16px; font-weight: 600; margin: 0 0 5px 0; padding-bottom: 6px; border-bottom: 2px solid #e2e8f0; text-align: center; }' +
                '.filter-showing-count { text-align: center; font-size: 12px; color: #013220; font-weight: 600; margin-bottom: 8px; }' +
                '.filter-cards-row { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 6px; margin-bottom: 8px; }' +
                '.filter-card { background: #F5F5F5; padding: 8px; border-radius: 6px; border-left: 4px solid #666; }' +
                '.filter-card-placeholder { }' +
                '.filter-subsection-header { font-size: 12px; font-weight: 600; color: #013220; margin-bottom: 8px; text-transform: uppercase; }' +
                '.filter-item { display: flex; align-items: center; gap: 8px; cursor: pointer; margin-bottom: 6px; }' +
                '.filter-item:last-of-type { margin-bottom: 0; }' +
                '.filter-item input[type="checkbox"] { width: 16px; height: 16px; min-width: 16px; min-height: 16px; cursor: pointer; accent-color: #013220; flex-shrink: 0; }' +
                '.filter-label { font-size: 13px; color: #1a2e1f; font-weight: 500; user-select: none; }' +
                '.filter-count { font-size: 11px; color: #6b7c72; margin-top: 4px; }' +
                /* Table controls */
                '.table-controls { margin: 20px 0; display: flex; gap: 10px; align-items: center; }' +
                '#searchBox { flex: 1; padding: 8px 12px; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 13px; }' +
                '#searchBox:focus { outline: none; border-color: #013220; box-shadow: 0 0 0 2px rgba(1,50,32,0.15); }' +
                '.action-btn-large { background: #013220; color: white; border: none; padding: 10px 24px; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 600; transition: background 0.15s; }' +
                '.action-btn-large:hover:not(:disabled) { background: #012618; }' +
                '.action-btn-large:disabled { background: #cbd5e1; cursor: not-allowed; }' +
                '.bulk-action-container { display: flex; gap: 8px; align-items: center; }' +
                '.bulk-action-select { padding: 8px 12px; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 13px; font-weight: 600; cursor: pointer; background: white; color: #013220; }' +
                '.bulk-action-select:focus { outline: none; border-color: #013220; box-shadow: 0 0 0 2px rgba(1,50,32,0.15); }' +
                '.bulk-action-select:disabled { opacity: 0.5; cursor: not-allowed; }' +
                /* Table styling */
                '.table-wrapper { margin: 20px 0; }' +
                '#dataTable { border-collapse: separate; border-spacing: 0; width: 100%; font-size: 12px; background: white; border: 1px solid #cbd5e1; }' +
                '#dataTable th { padding: 10px 8px; text-align: left; font-weight: 600; cursor: pointer; user-select: none; color: white; border: none; border-bottom: 2px solid #cbd5e1; position: -webkit-sticky; position: sticky; top: 0; z-index: 100; box-shadow: 0 2px 4px rgba(0,0,0,0.15); }' +
                '.th-checkbox { background: #013220; cursor: default !important; text-align: center; width: 40px; }' +
                '.th-actions { background: #013220; cursor: default !important; text-align: center; width: 50px; color: white; font-weight: 600; padding: 10px 8px; }' +
                '.th-slate { background: #013220; }' +
                '.th-slate:hover { background: #012618; }' +
                '.th-follow-up { width: 85px; }' +
                '.th-teal { background: #355E3B; }' +
                '.th-teal:hover { background: #2a4a2f; }' +
                '.sort-arrow { margin-left: 5px; opacity: 0.7; font-size: 10px; }' +
                '#dataTable td { border: none; padding: 8px 6px; vertical-align: middle; font-size: 12px; font-family: Arial, sans-serif; color: #1a2e1f; position: relative; z-index: 1; }' +
                '#dataTable tbody tr { border-bottom: 1px solid #d4e0d7; }' +
                '.col-checkbox { background: #E6EBE9; text-align: center; }' +
                '.col-actions { background: #E6EBE9; text-align: center; padding: 4px; }' +
                '.actions-dropdown { padding: 4px 6px; border: none; border-radius: 3px; font-size: 11px; cursor: pointer; background: transparent; color: #013220; font-weight: 600; width: 25px; appearance: none; -webkit-appearance: none; -moz-appearance: none; }' +
                '.actions-dropdown:hover { background: rgba(1, 50, 32, 0.05); }' +
                '.actions-dropdown:focus { outline: none; background: rgba(1, 50, 32, 0.1); }' +
                '.col-slate { background: #E6EBE9; max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }' +
                '.follow-up-cell { width: 85px; text-align: center; font-size: 11px; }' +
                '.col-teal { background: #EBF0EB; }' +
                '#dataTable tbody tr:hover td.col-checkbox { background: #CDD7D3; }' +
                '#dataTable tbody tr:hover td.col-actions { background: #CDD7D3; }' +
                '#dataTable tbody tr:hover td.col-slate { background: #CDD7D3; }' +
                '#dataTable tbody tr:hover td.col-teal { background: #D7E0D8; }' +
                '.amount { text-align: right; white-space: nowrap; color: #1a2e1f; }' +
                '.no-data { color: #6b7c72; font-style: italic; font-size: 12px; }' +
                '.transaction-link { color: #013220; text-decoration: none; font-weight: 600; font-size: 12px; }' +
                '.transaction-link:hover { text-decoration: underline; color: #355E3B; }' +
                '.customer-link { color: #2d3a33; text-decoration: none; font-size: 12px; }' +
                '.customer-link:hover { text-decoration: underline; color: #013220; }' +
                '.queued-cell { text-align: center; font-size: 16px; color: #355E3B; font-weight: bold; }' +
                '.queued-checkmark { color: #355E3B; }' +
                '.unqueue-x { color: #8B0000; font-size: 14px; margin-left: 6px; cursor: pointer; opacity: 0.6; font-weight: bold; transition: opacity 0.2s; }' +
                '.unqueue-x:hover { opacity: 1; }' +
                '.research-note-icon { margin-left: 4px; font-size: 12px; cursor: help; }' +
                '.research-notes-section { background: #FFF8DC; border: 1px solid #DAA520; border-radius: 4px; padding: 8px; margin-bottom: 12px; }' +
                '.research-notes-label { font-weight: 600; color: #B8860B; font-size: 11px; text-transform: uppercase; margin-bottom: 4px; }' +
                '.research-notes-value { color: #1a2e1f; font-size: 12px; line-height: 1.4; white-space: pre-wrap; word-wrap: break-word; }' +
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
                '.job-details-tooltip { display: none; position: fixed; bottom: 20px; right: 340px; background: white; border: 2px solid #013220; border-radius: 6px; padding: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.25); z-index: 100000; min-width: 300px; max-width: 450px; }' +
                '.job-details-tooltip.visible { display: block; }' +
                '.job-detail-section { margin-bottom: 12px; }' +
                '.job-detail-section:last-child { margin-bottom: 0; }' +
                '.job-detail-section.full-width { margin-bottom: 16px; padding-bottom: 12px; border-bottom: 1px solid #e2e8f0; }' +
                '.job-detail-columns { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }' +
                '.job-detail-column { display: flex; flex-direction: column; gap: 12px; }' +
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
                '.hidden { display: none !important; }' +
                /* Modal styles */
                '.modal-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 10000; display: flex; justify-content: center; align-items: center; }' +
                '.modal-content { background: white; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.3); width: 500px; max-width: 90%; }' +
                '.modal-header { background: #013220; color: white; padding: 15px 20px; font-size: 16px; font-weight: 600; border-radius: 8px 8px 0 0; }' +
                '.modal-body { padding: 20px; }' +
                '.modal-label { display: block; font-size: 13px; color: #1a2e1f; margin-bottom: 10px; font-weight: 500; }' +
                '.modal-textarea { width: 100%; padding: 12px; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 13px; font-family: Arial, sans-serif; resize: vertical; box-sizing: border-box; }' +
                '.modal-textarea:focus { outline: none; border-color: #013220; box-shadow: 0 0 0 2px rgba(1,50,32,0.15); }' +
                '.modal-input { width: 100%; padding: 10px; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 13px; font-family: Arial, sans-serif; box-sizing: border-box; }' +
                '.modal-input:focus { outline: none; border-color: #013220; box-shadow: 0 0 0 2px rgba(1,50,32,0.15); }' +
                '.modal-footer { padding: 15px 20px; border-top: 1px solid #e2e8f0; display: flex; justify-content: flex-end; gap: 10px; }' +
                '.modal-btn { padding: 10px 20px; border-radius: 4px; font-size: 14px; font-weight: 600; cursor: pointer; transition: all 0.15s; }' +
                '.modal-btn-cancel { background: white; color: #013220; border: 1px solid #cbd5e1; }' +
                '.modal-btn-cancel:hover { background: #f1f5f9; border-color: #013220; }' +
                '.modal-btn-save { background: #013220; color: white; border: none; }' +
                '.modal-btn-save:hover { background: #012618; }';
        }

        /**
         * Returns JavaScript for interactive features
         * @returns {string} JavaScript content
         */
        function getJavaScript() {
            return 'function filterTable() {' +
                '  applyFilters();' +
                '}' +
                'function applyFilters() {' +
                '  var input = document.getElementById("searchBox");' +
                '  var filter = input ? input.value.toUpperCase() : "";' +
                '  var oldShipDatesCheckbox = document.getElementById("filterOldShipDates");' +
                '  var futureShipDatesCheckbox = document.getElementById("filterFutureShipDates");' +
                '  var noShipDateCheckbox = document.getElementById("filterNoShipDate");' +
                '  var showOldShipDates = oldShipDatesCheckbox ? oldShipDatesCheckbox.checked : true;' +
                '  var showFutureShipDates = futureShipDatesCheckbox ? futureShipDatesCheckbox.checked : false;' +
                '  var showNoShipDate = noShipDateCheckbox ? noShipDateCheckbox.checked : true;' +
                '  var hasResearchNotesCheckbox = document.getElementById("filterHasResearchNotes");' +
                '  var noResearchNotesCheckbox = document.getElementById("filterNoResearchNotes");' +
                '  var showHasResearchNotes = hasResearchNotesCheckbox ? hasResearchNotesCheckbox.checked : true;' +
                '  var showNoResearchNotes = noResearchNotesCheckbox ? noResearchNotesCheckbox.checked : true;' +
                '  var jobScheduledCheckbox = document.getElementById("filterJobScheduled");' +
                '  var jobActiveCheckbox = document.getElementById("filterJobActive");' +
                '  var jobPausedCheckbox = document.getElementById("filterJobPaused");' +
                '  var jobCompletedCheckbox = document.getElementById("filterJobCompleted");' +
                '  var showJobScheduled = jobScheduledCheckbox ? jobScheduledCheckbox.checked : true;' +
                '  var showJobActive = jobActiveCheckbox ? jobActiveCheckbox.checked : true;' +
                '  var showJobPaused = jobPausedCheckbox ? jobPausedCheckbox.checked : true;' +
                '  var showJobCompleted = jobCompletedCheckbox ? jobCompletedCheckbox.checked : true;' +
                '  var warrantyNoneCheckbox = document.getElementById("filterWarrantyNone");' +
                '  var warrantyCBSICheckbox = document.getElementById("filterWarrantyCBSI");' +
                '  var warrantyCODCheckbox = document.getElementById("filterWarrantyCOD");' +
                '  var warrantyExtendedCheckbox = document.getElementById("filterWarrantyExtended");' +
                '  var warrantyMfgCheckbox = document.getElementById("filterWarrantyMfg");' +
                '  var warrantyShopCheckbox = document.getElementById("filterWarrantyShop");' +
                '  var showWarrantyNone = warrantyNoneCheckbox ? warrantyNoneCheckbox.checked : true;' +
                '  var showWarrantyCBSI = warrantyCBSICheckbox ? warrantyCBSICheckbox.checked : true;' +
                '  var showWarrantyCOD = warrantyCODCheckbox ? warrantyCODCheckbox.checked : true;' +
                '  var showWarrantyExtended = warrantyExtendedCheckbox ? warrantyExtendedCheckbox.checked : true;' +
                '  var showWarrantyMfg = warrantyMfgCheckbox ? warrantyMfgCheckbox.checked : true;' +
                '  var showWarrantyShop = warrantyShopCheckbox ? warrantyShopCheckbox.checked : true;' +
                '  var tbody = document.getElementById("reportTableBody");' +
                '  if (!tbody) return;' +
                '  var tr = tbody.children;' +
                '  var today = new Date();' +
                '  today.setHours(0, 0, 0, 0);' +
                '  var visibleCount = 0;' +
                '  var totalCount = tr.length;' +
                '  for (var i = 0; i < tr.length; i++) {' +
                '    var row = tr[i];' +
                '    var showRow = true;' +
                '    var txtValue = row.textContent || row.innerText;' +
                '    if (filter && txtValue.toUpperCase().indexOf(filter) === -1) {' +
                '      showRow = false;' +
                '    }' +
                '    if (showRow) {' +
                '      var researchNotes = row.getAttribute("data-research-notes") || "";' +
                '      var hasNotes = researchNotes.trim().length > 0;' +
                '      if (hasNotes && !showHasResearchNotes) {' +
                '        showRow = false;' +
                '      } else if (!hasNotes && !showNoResearchNotes) {' +
                '        showRow = false;' +
                '      }' +
                '    }' +
                '    if (showRow) {' +
                '      var shipDateStr = row.getAttribute("data-ship-date");' +
                '      if (!shipDateStr) {' +
                '        showRow = showNoShipDate;' +
                '      } else {' +
                '        var shipDate = new Date(shipDateStr);' +
                '        shipDate.setHours(0, 0, 0, 0);' +
                '        if (shipDate >= today) {' +
                '          showRow = showFutureShipDates;' +
                '        } else {' +
                '          showRow = showOldShipDates;' +
                '        }' +
                '      }' +
                '    }' +
                '    if (showRow) {' +
                '      var jobState = (row.getAttribute("data-job-state") || "").toLowerCase();' +
                '      var jobStateMatch = false;' +
                '      if (showJobScheduled && jobState.indexOf("scheduled") >= 0) jobStateMatch = true;' +
                '      if (showJobActive && jobState.indexOf("active") >= 0) jobStateMatch = true;' +
                '      if (showJobPaused && jobState.indexOf("paused") >= 0) jobStateMatch = true;' +
                '      if (showJobCompleted && jobState.indexOf("completed") >= 0) jobStateMatch = true;' +
                '      if (!jobState && (showJobScheduled || showJobActive || showJobPaused || showJobCompleted)) jobStateMatch = true;' +
                '      showRow = jobStateMatch;' +
                '    }' +
                '    if (showRow) {' +
                '      var warrantyType = (row.getAttribute("data-warranty-type") || "").toLowerCase();' +
                '      var warrantyMatch = false;' +
                '      if (showWarrantyCBSI && warrantyType === "cbsi") warrantyMatch = true;' +
                '      else if (showWarrantyCOD && warrantyType === "cash on delivery") warrantyMatch = true;' +
                '      else if (showWarrantyExtended && warrantyType === "extended warranty") warrantyMatch = true;' +
                '      else if (showWarrantyMfg && warrantyType === "manufacturer warranty term") warrantyMatch = true;' +
                '      else if (showWarrantyShop && warrantyType === "shop ticket") warrantyMatch = true;' +
                '      else if (showWarrantyNone && (warrantyType === "" || warrantyType === "-- please select --")) warrantyMatch = true;' +
                '      showRow = warrantyMatch;' +
                '    }' +
                '    if (showRow) {' +
                '      row.style.display = "";' +
                '      visibleCount++;' +
                '    } else {' +
                '      row.style.display = "none";' +
                '    }' +
                '  }' +
                '  var filterCountEl = document.getElementById("filterShipDateCount");' +
                '  if (filterCountEl) {' +
                '    filterCountEl.textContent = "Showing: " + visibleCount + " / " + totalCount;' +
                '  }' +
                '  updateSelectedSummary();' +
                '}' +
                'function handleUnqueue(soId, event) {' +
                '  event.stopPropagation();' +
                '  if (!confirm(\"Remove this Sales Order from the Bill & Write-Off queue?\\\\n\\\\nThis will clear the queue date and allow it to be processed later.\")) {' +
                '    return;' +
                '  }' +
                '  showLoading();' +
                '  var xhr = new XMLHttpRequest();' +
                '  xhr.open(\"POST\", SUITELET_URL, true);' +
                '  xhr.setRequestHeader(\"Content-Type\", \"application/x-www-form-urlencoded\");' +
                '  xhr.onreadystatechange = function() {' +
                '    if (xhr.readyState === 4) {' +
                '      hideLoading();' +
                '      try {' +
                '        var resp = JSON.parse(xhr.responseText);' +
                '        if (resp.success) {' +
                '          alert(resp.message);' +
                '          var queuedCell = document.getElementById(\"queued-cell-\" + soId);' +
                '          if (queuedCell) {' +
                '            queuedCell.innerHTML = \"\";' +
                '          }' +
                '          var checkbox = document.querySelector(\".so-checkbox[value=\\\\\\\"\" + soId + \"\\\\\\\"]\");' +
                '          if (checkbox) {' +
                '            checkbox.disabled = false;' +
                '            checkbox.checked = false;' +
                '          }' +
                '          var queuedTotalEl = document.getElementById(\"queuedTotal\");' +
                '          var queuedTotalLinesEl = document.getElementById(\"queuedTotalLines\");' +
                '          var queuedTotalAmountEl = document.getElementById(\"queuedTotalAmount\");' +
                '          var row = queuedCell ? queuedCell.closest(\"tr\") : null;' +
                '          if (row) {' +
                '            var unbilledLines = parseInt(row.getAttribute(\"data-unbilled-lines\") || 0);' +
                '            var unbilledAmount = parseFloat(row.getAttribute(\"data-unbilled-amount\") || 0);' +
                '            if (queuedTotalEl) queuedTotalEl.textContent = Math.max(0, parseInt(queuedTotalEl.textContent || 0) - 1);' +
                '            if (queuedTotalLinesEl) queuedTotalLinesEl.textContent = Math.max(0, parseInt(queuedTotalLinesEl.textContent || 0) - unbilledLines);' +
                '            if (queuedTotalAmountEl) {' +
                '              var currentAmt = parseFloat(queuedTotalAmountEl.textContent.replace(/[^0-9.-]/g, \"\") || 0);' +
                '              queuedTotalAmountEl.textContent = \"$\" + Math.max(0, currentAmt - unbilledAmount).toFixed(2).replace(/\\\\B(?=(\\\\d{3})+(?!\\\\d))/g, \",\");' +
                '            }' +
                '          }' +
                '        } else {' +
                '          alert(\"Error: \" + resp.message);' +
                '        }' +
                '      } catch (e) {' +
                '        alert(\"Error processing response: \" + e.toString());' +
                '      }' +
                '    }' +
                '  };' +
                '  xhr.send(\"action=unqueue&soId=\" + soId);' +
                '}' +
                'var sortDir = {};' +
                'function sortTable(n) {' +
                '  var table = document.getElementById("dataTable");' +
                '  var tbody = document.getElementById("reportTableBody");' +
                '  var rows = Array.from(tbody.rows);' +
                '  var numericCols = [13, 14];' +
                '  var dateCols = [3, 9, 10, 11];' +
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
                '    var row = checkboxes[i].closest("tr");' +
                '    if (row && row.style.display !== "none") {' +
                '      checkboxes[i].checked = checkbox.checked;' +
                '    }' +
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
                '  var bulkActionSelect = document.getElementById("bulkActionSelect");' +
                '  var bulkActionBtn = document.getElementById("bulkActionBtn");' +
                '  if (selectedCountEl) selectedCountEl.textContent = selectedCount;' +
                '  if (selectedLinesEl) selectedLinesEl.textContent = selectedLines;' +
                '  if (selectedAmountEl) selectedAmountEl.textContent = "$" + selectedAmount.toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",");' +
                '  if (bulkActionSelect) {' +
                '    bulkActionSelect.disabled = (selectedCount === 0);' +
                '  }' +
                '  if (bulkActionBtn) {' +
                '    bulkActionBtn.disabled = (selectedCount === 0);' +
                '  }' +
                '}' +
                'function executeBulkAction() {' +
                '  var bulkActionSelect = document.getElementById("bulkActionSelect");' +
                '  var action = bulkActionSelect ? bulkActionSelect.value : "";' +
                '  if (!action) {' +
                '    alert("Please select a bulk action from the dropdown.");' +
                '    return;' +
                '  }' +
                '  var checkboxes = document.querySelectorAll(".so-checkbox:checked");' +
                '  if (checkboxes.length === 0) {' +
                '    alert("Please select at least one Sales Order.");' +
                '    return;' +
                '  }' +
                '  var soIds = [];' +
                '  for (var i = 0; i < checkboxes.length; i++) {' +
                '    soIds.push(checkboxes[i].value);' +
                '  }' +
                '  var actionLabels = {' +
                '    "queue": "Queue for Bill & Write-Off",' +
                '    "close": "Close (Cancel)",' +
                '    "auto-bill": "Auto-Bill (Invoice)",' +
                '    "cbsi-bill-je": "CBSI (Bill and JE)"' +
                '  };' +
                '  var actionLabel = actionLabels[action] || action;' +
                '  var confirmMsg = actionLabel + " " + soIds.length + " Sales Order(s)?\\n\\n";' +
                '  if (action === "queue") {' +
                '    confirmMsg += "This will mark them as queued for future processing.";' +
                '  } else if (action === "close") {' +
                '    confirmMsg += "This will close all selected orders and cannot be easily undone.";' +
                '  } else if (action === "auto-bill") {' +
                '    confirmMsg += "This will create invoices for all selected orders.";' +
                '  } else if (action === "cbsi-bill-je") {' +
                '    confirmMsg += "This will:\\n1. Create invoices with entity CBSI\\n2. Create write-off JEs\\n3. Apply JEs to invoices\\n\\nThis is the most complex operation and may hit governance limits with many records.";' +
                '  }' +
                '  if (!confirm(confirmMsg)) {' +
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
                '          handleBulkActionResponse(action, resp);' +
                '        } else {' +
                '          alert("Error: " + resp.message);' +
                '        }' +
                '      } catch (e) {' +
                '        alert("Error processing response: " + e.toString());' +
                '      }' +
                '    }' +
                '  };' +
                '  xhr.send("selectedSOIds=" + soIds.join(",") + "&bulkAction=" + action);' +
                '}' +
                'function handleBulkActionResponse(action, resp) {' +
                '  var processedIds = resp.processedIds || resp.queuedIds || [];' +
                '  var failedIds = resp.failedIds || [];' +
                '  if (action === "queue") {' +
                '    var queuedLines = 0;' +
                '    var queuedAmount = 0;' +
                '    for (var i = 0; i < processedIds.length; i++) {' +
                '      var queuedCell = document.getElementById("queued-cell-" + processedIds[i]);' +
                '      if (queuedCell) {' +
                '        queuedCell.innerHTML = "<span class=\\"queued-checkmark\\">✓</span><span class=\\"unqueue-x\\" onclick=\\"handleUnqueue(" + processedIds[i] + ", event)\\" title=\\"Remove from queue\\">✕</span>";' +
                '      }' +
                '      var checkbox = document.querySelector(".so-checkbox[value=\\\"" + processedIds[i] + "\\\"]");' +
                '      if (checkbox) {' +
                '        var row = checkbox.closest("tr");' +
                '        if (row) {' +
                '          queuedLines += parseInt(row.getAttribute("data-unbilled-lines") || 0);' +
                '          queuedAmount += parseFloat(row.getAttribute("data-unbilled-amount") || 0);' +
                '        }' +
                '        checkbox.checked = resp.governanceStopped ? true : false;' +
                '        checkbox.disabled = resp.governanceStopped ? false : true;' +
                '      }' +
                '    }' +
                '    var queuedTotalEl = document.getElementById("queuedTotal");' +
                '    var queuedTotalLinesEl = document.getElementById("queuedTotalLines");' +
                '    var queuedTotalAmountEl = document.getElementById("queuedTotalAmount");' +
                '    if (queuedTotalEl) queuedTotalEl.textContent = parseInt(queuedTotalEl.textContent || 0) + processedIds.length;' +
                '    if (queuedTotalLinesEl) queuedTotalLinesEl.textContent = parseInt(queuedTotalLinesEl.textContent || 0) + queuedLines;' +
                '    if (queuedTotalAmountEl) {' +
                '      var currentAmt = parseFloat(queuedTotalAmountEl.textContent.replace(/[^0-9.-]/g, "") || 0);' +
                '      queuedTotalAmountEl.textContent = "$" + (currentAmt + queuedAmount).toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",");' +
                '    }' +
                '  } else if (action === "close" || action === "auto-bill" || action === "cbsi-bill-je") {' +
                '    for (var i = 0; i < processedIds.length; i++) {' +
                '      var checkbox = document.querySelector(".so-checkbox[value=\\\"" + processedIds[i] + "\\\"]");' +
                '      if (checkbox) {' +
                '        var row = checkbox.closest("tr");' +
                '        if (row && !resp.governanceStopped) {' +
                '          row.style.display = "none";' +
                '        } else if (row && resp.governanceStopped) {' +
                '          checkbox.checked = false;' +
                '        }' +
                '      }' +
                '    }' +
                '  }' +
                '  updateSelectedSummary();' +
                '  alert(resp.message);' +
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
                '            var queuedTotal = document.getElementById("queuedTotal");' +
                '            var queuedTotalLines = document.getElementById("queuedTotalLines");' +
                '            var queuedTotalAmount = document.getElementById("queuedTotalAmount");' +
                '            var reportTableBody = document.getElementById("reportTableBody");' +
                '            var reportContent = document.getElementById("reportContent");' +
                '            if (summaryTotal) summaryTotal.textContent = resp.summaryTotal;' +
                '            if (summaryTotalLines) summaryTotalLines.textContent = resp.summaryTotalLines;' +
                '            if (summaryTotalAmount) summaryTotalAmount.textContent = "$" + parseFloat(resp.summaryTotalAmount).toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",");' +
                '            if (queuedTotal) queuedTotal.textContent = resp.queuedTotal;' +
                '            if (queuedTotalLines) queuedTotalLines.textContent = resp.queuedTotalLines;' +
                '            if (queuedTotalAmount) queuedTotalAmount.textContent = "$" + parseFloat(resp.queuedTotalAmount).toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",");' +
                '            if (reportTableBody) reportTableBody.innerHTML = resp.tableBodyHtml;' +
                '            if (reportContent) {' +
                '              reportContent.className = "";' +
                '              reportContent.style.display = "block";' +
                '            }' +
                '            applyFilters();' +
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
                'function handleUnqueue(soId, event) {' +
                '  event.stopPropagation();' +
                '  if (!confirm("Remove this Sales Order from the Bill & Write-Off queue?\\n\\nThis will clear the queue date and allow it to be processed later.")) {' +
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
                '          alert(resp.message);' +
                '          var queuedCell = document.getElementById("queued-cell-" + soId);' +
                '          if (queuedCell) {' +
                '            queuedCell.innerHTML = "";' +
                '          }' +
                '          var checkbox = document.querySelector(".so-checkbox[value=\\\"" + soId + "\\\"]");' +
                '          if (checkbox) {' +
                '            checkbox.disabled = false;' +
                '            checkbox.checked = false;' +
                '          }' +
                '          var queuedTotalEl = document.getElementById("queuedTotal");' +
                '          var queuedTotalLinesEl = document.getElementById("queuedTotalLines");' +
                '          var queuedTotalAmountEl = document.getElementById("queuedTotalAmount");' +
                '          var row = queuedCell ? queuedCell.closest("tr") : null;' +
                '          if (row) {' +
                '            var unbilledLines = parseInt(row.getAttribute("data-unbilled-lines") || 0);' +
                '            var unbilledAmount = parseFloat(row.getAttribute("data-unbilled-amount") || 0);' +
                '            if (queuedTotalEl) queuedTotalEl.textContent = Math.max(0, parseInt(queuedTotalEl.textContent || 0) - 1);' +
                '            if (queuedTotalLinesEl) queuedTotalLinesEl.textContent = Math.max(0, parseInt(queuedTotalLinesEl.textContent || 0) - unbilledLines);' +
                '            if (queuedTotalAmountEl) {' +
                '              var currentAmt = parseFloat(queuedTotalAmountEl.textContent.replace(/[^0-9.-]/g, "") || 0);' +
                '              queuedTotalAmountEl.textContent = "$" + Math.max(0, currentAmt - unbilledAmount).toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",");' +
                '            }' +
                '          }' +
                '        } else {' +
                '          alert("Error: " + resp.message);' +
                '        }' +
                '      } catch (e) {' +
                '        alert("Error processing response: " + e.toString());' +
                '      }' +
                '    }' +
                '  };' +
                '  xhr.send("action=unqueue&soId=" + soId);' +
                '}' +
                'function handleUnqueue(soId, event) {' +
                '  event.stopPropagation();' +
                '  if (!confirm("Remove this Sales Order from the Bill & Write-Off queue?\\n\\nThis will clear the queue date and allow it to be processed later.")) {' +
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
                '          alert(resp.message);' +
                '          var queuedCell = document.getElementById("queued-cell-" + soId);' +
                '          if (queuedCell) {' +
                '            queuedCell.innerHTML = "";' +
                '          }' +
                '          var checkbox = document.querySelector(".so-checkbox[value=\\\"" + soId + "\\\"]");' +
                '          if (checkbox) {' +
                '            checkbox.disabled = false;' +
                '            checkbox.checked = false;' +
                '          }' +
                '          var queuedTotalEl = document.getElementById("queuedTotal");' +
                '          var queuedTotalLinesEl = document.getElementById("queuedTotalLines");' +
                '          var queuedTotalAmountEl = document.getElementById("queuedTotalAmount");' +
                '          var row = queuedCell ? queuedCell.closest("tr") : null;' +
                '          if (row) {' +
                '            var unbilledLines = parseInt(row.getAttribute("data-unbilled-lines") || 0);' +
                '            var unbilledAmount = parseFloat(row.getAttribute("data-unbilled-amount") || 0);' +
                '            if (queuedTotalEl) queuedTotalEl.textContent = Math.max(0, parseInt(queuedTotalEl.textContent || 0) - 1);' +
                '            if (queuedTotalLinesEl) queuedTotalLinesEl.textContent = Math.max(0, parseInt(queuedTotalLinesEl.textContent || 0) - unbilledLines);' +
                '            if (queuedTotalAmountEl) {' +
                '              var currentAmt = parseFloat(queuedTotalAmountEl.textContent.replace(/[^0-9.-]/g, "") || 0);' +
                '              queuedTotalAmountEl.textContent = "$" + Math.max(0, currentAmt - unbilledAmount).toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",");' +
                '            }' +
                '          }' +
                '        } else {' +
                '          alert("Error: " + resp.message);' +
                '        }' +
                '      } catch (e) {' +
                '        alert("Error processing response: " + e.toString());' +
                '      }' +
                '    }' +
                '  };' +
                '  xhr.send("action=unqueue&soId=" + soId);' +
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
                '  var researchNotes = row.getAttribute("data-research-notes");' +
                '  var followUpDate = row.getAttribute("data-follow-up-date");' +
                '  var jobDetails = row.getAttribute("data-job-details");' +
                '  var billingCompletedBy = row.getAttribute("data-billing-completed-by");' +
                '  var jobState = row.getAttribute("data-job-state");' +
                '  var partsStatus = row.getAttribute("data-parts-status");' +
                '  var scheduledDate = row.getAttribute("data-scheduled-date");' +
                '  var jobStarted = row.getAttribute("data-job-started");' +
                '  var jobCompleted = row.getAttribute("data-job-completed");' +
                '  var html = "";' +
                '  if (researchNotes) {' +
                '    html += "<div class=\\"research-notes-section\\"><div class=\\"research-notes-label\\">Research Notes:</div><div class=\\"research-notes-value\\">" + researchNotes.replace(/&lt;br&gt;/gi, "<br>") + "</div></div>";' +
                '  }' +
                '  if (followUpDate) {' +
                '    var formattedFollowUp = "";' +
                '    if (followUpDate.match(/^\\d{4}-\\d{2}-\\d{2}$/)) {' +
                '      var parts = followUpDate.split("-");' +
                '      var d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));' +
                '      formattedFollowUp = (d.getMonth() + 1) + "/" + d.getDate() + "/" + d.getFullYear();' +
                '    } else {' +
                '      formattedFollowUp = new Date(followUpDate).toLocaleDateString();' +
                '    }' +
                '    html += "<div class=\\"research-notes-section\\"><div class=\\"research-notes-label\\">Follow Up Date:</div><div class=\\"research-notes-value\\">" + formattedFollowUp + "</div></div>";' +
                '  }' +
                '  var unescapedDetails = jobDetails ? jobDetails.replace(/&lt;br&gt;/gi, "<br>") : "None";' +
                '  html += "<div class=\\"job-detail-section full-width\\"><div class=\\"job-detail-label\\">Job Details:</div><div class=\\"job-detail-value\\">" + unescapedDetails + "</div></div>";' +
                '  html += "<div class=\\"job-detail-columns\\">";' +
                '  html += "<div class=\\"job-detail-column\\">";' +
                '  html += "<div class=\\"job-detail-section\\"><div class=\\"job-detail-label\\">Billing Completed By:</div><div class=\\"job-detail-value\\">" + (billingCompletedBy || "None") + "</div></div>";' +
                '  html += "<div class=\\"job-detail-section\\"><div class=\\"job-detail-label\\">Job State:</div><div class=\\"job-detail-value\\">" + (jobState || "None") + "</div></div>";' +
                '  html += "<div class=\\"job-detail-section\\"><div class=\\"job-detail-label\\">Parts Status:</div><div class=\\"job-detail-value\\">" + (partsStatus || "None") + "</div></div>";' +
                '  html += "</div>";' +
                '  html += "<div class=\\"job-detail-column\\">";' +
                '  var formattedScheduled = scheduledDate ? new Date(scheduledDate).toLocaleDateString() : "None";' +
                '  html += "<div class=\\"job-detail-section\\"><div class=\\"job-detail-label\\">Scheduled Date:</div><div class=\\"job-detail-value\\">" + formattedScheduled + "</div></div>";' +
                '  var formattedStarted = jobStarted ? new Date(jobStarted).toLocaleDateString() : "None";' +
                '  html += "<div class=\\"job-detail-section\\"><div class=\\"job-detail-label\\">Job Started:</div><div class=\\"job-detail-value\\">" + formattedStarted + "</div></div>";' +
                '  var formattedCompleted = jobCompleted ? new Date(jobCompleted).toLocaleDateString() : "None";' +
                '  html += "<div class=\\"job-detail-section\\"><div class=\\"job-detail-label\\">Job Completed:</div><div class=\\"job-detail-value\\">" + formattedCompleted + "</div></div>";' +
                '  html += "</div>";' +
                '  html += "</div>";' +
                '  tooltipContent.innerHTML = html;' +
                '  tooltip.className = "job-details-tooltip visible";' +
                '}' +
                'function hideJobDetailsTooltip() {' +
                '  var tooltip = document.getElementById("jobDetailsTooltip");' +
                '  if (tooltip) tooltip.className = "job-details-tooltip";' +
                '}' +
                'function handleAction(selectElement, soId) {' +
                '  var action = selectElement.value;' +
                '  if (!action) return;' +
                '  selectElement.value = "";' +
                '  if (action === "queue") {' +
                '    if (!confirm("Queue this Sales Order for Bill & Write-Off?\\n\\nThis will mark it for processing by the scheduled script.")) {' +
                '      return;' +
                '    }' +
                '    showLoading();' +
                '    var xhr = new XMLHttpRequest();' +
                '    xhr.open("POST", SUITELET_URL, true);' +
                '    xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");' +
                '    xhr.onreadystatechange = function() {' +
                '      if (xhr.readyState === 4) {' +
                '        hideLoading();' +
                '        try {' +
                '          var resp = JSON.parse(xhr.responseText);' +
                '          if (resp.success) {' +
                '            alert(resp.message);' +
                '            var queuedCell = document.getElementById("queued-cell-" + soId);' +
                '            if (queuedCell) {' +
                '              queuedCell.innerHTML = "<span class=\\"queued-checkmark\\">✓</span><span class=\\"unqueue-x\\" onclick=\\"handleUnqueue(" + soId + ", event)\\" title=\\"Remove from queue\\">✕</span>";' +
                '            }' +
                '            var checkbox = document.querySelector(".so-checkbox[value=\\\"" + soId + "\\\"]");' +
                '            if (checkbox) {' +
                '              checkbox.checked = false;' +
                '              checkbox.disabled = true;' +
                '            }' +
                '            var queuedTotalEl = document.getElementById("queuedTotal");' +
                '            var queuedTotalLinesEl = document.getElementById("queuedTotalLines");' +
                '            var queuedTotalAmountEl = document.getElementById("queuedTotalAmount");' +
                '            var row = selectElement.closest("tr");' +
                '            if (row) {' +
                '              var unbilledLines = parseInt(row.getAttribute("data-unbilled-lines") || 0);' +
                '              var unbilledAmount = parseFloat(row.getAttribute("data-unbilled-amount") || 0);' +
                '              if (queuedTotalEl) queuedTotalEl.textContent = parseInt(queuedTotalEl.textContent || 0) + 1;' +
                '              if (queuedTotalLinesEl) queuedTotalLinesEl.textContent = parseInt(queuedTotalLinesEl.textContent || 0) + unbilledLines;' +
                '              if (queuedTotalAmountEl) {' +
                '                var currentAmt = parseFloat(queuedTotalAmountEl.textContent.replace(/[^0-9.-]/g, "") || 0);' +
                '                queuedTotalAmountEl.textContent = "$" + (currentAmt + unbilledAmount).toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",");' +
                '              }' +
                '            }' +
                '          } else {' +
                '            alert("Error: " + resp.message);' +
                '          }' +
                '        } catch (e) {' +
                '          alert("Error processing response: " + e.toString());' +
                '        }' +
                '      }' +
                '    };' +
                '    xhr.send("action=queue&soId=" + soId);' +
                '  } else if (action === "close") {' +
                '    if (!confirm("Close (Cancel) this Sales Order?\\n\\nThis will set the status to Closed and cannot be easily undone.")) {' +
                '      return;' +
                '    }' +
                '    showLoading();' +
                '    var xhr = new XMLHttpRequest();' +
                '    xhr.open("POST", SUITELET_URL, true);' +
                '    xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");' +
                '    xhr.onreadystatechange = function() {' +
                '      if (xhr.readyState === 4) {' +
                '        hideLoading();' +
                '        try {' +
                '          var resp = JSON.parse(xhr.responseText);' +
                '          if (resp.success) {' +
                '            alert(resp.message);' +
                '            var row = selectElement.closest("tr");' +
                '            if (row) row.style.display = "none";' +
                '          } else {' +
                '            alert("Error: " + resp.message);' +
                '          }' +
                '        } catch (e) {' +
                '          alert("Error processing response: " + e.toString());' +
                '        }' +
                '      }' +
                '    };' +
                '    xhr.send("action=close&soId=" + soId);' +
                '  } else if (action === "bill") {' +
                '    var invoiceUrl = "/app/accounting/transactions/custinvc.nl?id=" + soId + "&e=T&transform=salesord&billremaining=T&memdoc=0&whence=";' +
                '    window.open(invoiceUrl, "_blank");' +
                '  } else if (action === "auto-bill") {' +
                '    if (!confirm("Automatically create invoice for this Sales Order?\\n\\nThis will transform and save the invoice immediately.")) {' +
                '      return;' +
                '    }' +
                '    showLoading();' +
                '    var xhr = new XMLHttpRequest();' +
                '    xhr.open("POST", SUITELET_URL, true);' +
                '    xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");' +
                '    xhr.onreadystatechange = function() {' +
                '      if (xhr.readyState === 4) {' +
                '        hideLoading();' +
                '        try {' +
                '          var resp = JSON.parse(xhr.responseText);' +
                '          if (resp.success) {' +
                '            alert(resp.message + " Invoice: " + resp.invoiceTranid);' +
                '            var row = selectElement.closest("tr");' +
                '            if (row) row.style.display = "none";' +
                '          } else {' +
                '            alert("Error: " + resp.message);' +
                '          }' +
                '        } catch (e) {' +
                '          alert("Error processing response: " + e.toString());' +
                '        }' +
                '      }' +
                '    };' +
                '    xhr.send("action=auto-bill&soId=" + soId);' +
                '  } else if (action === "cbsi-bill-je") {' +
                '    if (!confirm("Process CBSI Bill and JE for this Sales Order?\\n\\nThis will:\\n1. Create invoice with entity CBSI (335)\\n2. Create write-off journal entry\\n3. Apply JE to invoice\\n\\nThis cannot be easily undone.")) {' +
                '      return;' +
                '    }' +
                '    showLoading();' +
                '    var xhr = new XMLHttpRequest();' +
                '    xhr.open("POST", SUITELET_URL, true);' +
                '    xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");' +
                '    xhr.onreadystatechange = function() {' +
                '      if (xhr.readyState === 4) {' +
                '        hideLoading();' +
                '        try {' +
                '          var resp = JSON.parse(xhr.responseText);' +
                '          if (resp.success) {' +
                '            alert(resp.message + "\\n\\nInvoice: " + resp.invoiceTranid + "\\nJE: " + resp.jeTranid + "\\nAmount: $" + resp.amount.toFixed(2));' +
                '            var row = selectElement.closest("tr");' +
                '            if (row) row.style.display = "none";' +
                '          } else {' +
                '            alert("Error: " + resp.message);' +
                '          }' +
                '        } catch (e) {' +
                '          alert("Error processing response: " + e.toString());' +
                '        }' +
                '      }' +
                '    };' +
                '    xhr.send("action=cbsi-bill-je&soId=" + soId);' +
                '  } else if (action === "add-note") {' +
                '    var row = selectElement.closest("tr");' +
                '    var existingNote = row ? row.getAttribute("data-research-notes") || "" : "";' +
                '    openResearchNoteModal(soId, existingNote, row);' +
                '  }' +
                '}' +
                'var currentNoteSOId = null;' +
                'var currentNoteRow = null;' +
                'function openResearchNoteModal(soId, existingNote, row) {' +
                '  currentNoteSOId = soId;' +
                '  currentNoteRow = row;' +
                '  var modal = document.getElementById("researchNoteModal");' +
                '  var textarea = document.getElementById("researchNoteInput");' +
                '  var dateInput = document.getElementById("followUpDateInput");' +
                '  if (modal && textarea) {' +
                '    textarea.value = existingNote;' +
                '    if (dateInput) {' +
                '      var existingDate = row ? (row.getAttribute("data-follow-up-date") || "") : "";' +
                '      dateInput.value = existingDate;' +
                '    }' +
                '    modal.style.display = "flex";' +
                '    textarea.focus();' +
                '  }' +
                '}' +
                'function closeResearchNoteModal() {' +
                '  var modal = document.getElementById("researchNoteModal");' +
                '  if (modal) modal.style.display = "none";' +
                '  currentNoteSOId = null;' +
                '  currentNoteRow = null;' +
                '}' +
                'function saveResearchNote() {' +
                '  var textarea = document.getElementById("researchNoteInput");' +
                '  var dateInput = document.getElementById("followUpDateInput");' +
                '  var newNote = textarea ? textarea.value : "";' +
                '  var followUpDate = dateInput ? dateInput.value : "";' +
                '  var soId = currentNoteSOId;' +
                '  var row = currentNoteRow;' +
                '  closeResearchNoteModal();' +
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
                '          alert(resp.message);' +
                '          if (row) {' +
                '            row.setAttribute("data-research-notes", newNote);' +
                '            row.setAttribute("data-follow-up-date", followUpDate);' +
                '            var noteIcon = document.getElementById("note-icon-" + soId);' +
                '            if (noteIcon) {' +
                '              noteIcon.style.display = newNote ? "inline" : "none";' +
                '            }' +
                '            var followUpCell = document.getElementById("follow-up-cell-" + soId);' +
                '            if (followUpCell) {' +
                '              var formattedDate = "";' +
                '              if (followUpDate && followUpDate.match(/^\\d{4}-\\d{2}-\\d{2}$/)) {' +
                '                var parts = followUpDate.split("-");' +
                '                var d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));' +
                '                formattedDate = (d.getMonth() + 1) + "/" + d.getDate() + "/" + d.getFullYear();' +
                '              }' +
                '              followUpCell.textContent = formattedDate;' +
                '            }' +
                '          }' +
                '          applyFilters();' +
                '        } else {' +
                '          alert("Error: " + resp.message);' +
                '        }' +
                '      } catch (e) {' +
                '        alert("Error processing response: " + e.toString());' +
                '      }' +
                '    }' +
                '  };' +
                '  xhr.send("action=add-note&soId=" + soId + "&note=" + encodeURIComponent(newNote) + "&followUpDate=" + encodeURIComponent(followUpDate));' +
                '}';
        }

        return {
            onRequest: onRequest
        };
    });