/*global Logger, SpreadsheetApp, HtmlService, Analytics */

/*
 *
 * Utility/helper functions
 *
 */

function replaceSparseInArrayWithDefault(array, defaultValue, length) {
    var i;
    for (i = 0; i < length; i++) {
        if (!(i in array)) {
            array[i] = defaultValue;
        }
    }

    return array;
}

function validateCallback(callback) {
    if (typeof callback !== 'function') {
        throw new Error('Callback must be a function');
    }
}

/*
 *
 * SHEET
 *
 */

var Sheet = function (name, data) {
    this.workbook = SpreadsheetApp.getActiveSpreadsheet();
    this.name = name;
    this.data = data;
    this.sheet = this.workbook.getSheetByName(this.name) || this.workbook.insertSheet(this.name);
};

Sheet.prototype = {

    getHeaderRow: function () {
        var colors = {
            blue: '#4d90fe',
            green: '#0fc357',
            purple: '#c27ba0',
            yellow: '#e7fe2b'
        };

        var cells;

        switch (this.name) {

        case 'Filters - View Level':

            cells = [{
                name: 'Account Name'
            }, {
                name: 'Account Id'
            }, {
                name: 'Property Name'
            }, {
                name: 'Property Id'
            }, {
                name: 'View Name'
            }, {
                name: 'View Id'
            }, {
                name: 'Filter Name'
            }, {
                name: 'Filter Id'
            }];

            break;

        case 'Views':

            cells = [{
                name: 'Account Name'
            }, {
                name: 'Account Id'
            }, {
                name: 'Property Name'
            }, {
                name: 'Property Id'
            }, {
                name: 'View Name'
            }, {
                name: 'View Id'
            }, {
                name: 'View Website URL'
            }, {
                name: 'View Timezone'
            }, {
                name: 'View Default Page'
            }, {
                name: 'View Exclude Query Parameters'
            }, {
                name: 'View Currency'
            }, {
                name: 'View Site Search Query Parameters'
            }, {
                name: 'View Strip Site Search Query Parameters'
            }, {
                name: 'View Site Search Category Parameters',
            }, {
                name: 'View Strip Site Search Category Parameters'
            }];

            break;

        case 'Filters - Account Level':

            cells = [{
                name: 'Account Name'
            }, {
                name: 'Account Id'
            }, {
                name: 'Filter Name'
            }, {
                name: 'Filter Id'
            }, {
                name: 'Filter Type'
            }, {
                name: 'Filter Field'
            }, {
                name: 'Filter MatchType'
            }, {
                name: 'Filter ExpressionValue'
            }, {
                name: 'Filter CaseSensitive'
            }, {
                name: 'Filter SearchString'
            }, {
                name: 'Filter ReplaceString'
            }, {
                name: 'Filter FieldA'
            }, {
                name: 'Filter ExtractA'
            }, {
                name: 'Filter FieldB'
            }, {
                name: 'Filter ExtractB'
            }, {
                name: 'Filter OutputToField'
            }, {
                name: 'Filter OutputConstructor'
            }, {
                name: 'Filter FieldARequired'
            }, {
                name: 'Filter FieldBRequired'
            }, {
                name: 'Filter OverrideOutputField'
            }, {
                name: 'Filter CaseSensitive'
            }];

            break;

        case 'Goals':

            cells = [{
                name: 'Account Name'
            }, {
                name: 'Account Id'
            }, {
                name: 'Property Name'
            }, {
                name: 'Property Id'
            }, {
                name: 'View Name'
            }, {
                name: 'View Id'
            }, {
                name: 'Goal Name'
            }, {
                name: 'Goal Id'
            }, {
                name: 'Goal Type'
            }, {
                name: 'Goal Active'
            }, {
                name: 'Goal Value'
            }, {
                name: 'Goal Detail URL',
                color: colors.green
            }, {
                name: 'Goal Detail CaseSensitive',
                color: colors.green
            }, {
                name: 'Goal Detail MatchType',
                color: colors.green
            }, {
                name: 'Goal Detail FirstStepRequired',
                color: colors.green
            }, {
                name: 'Goal Detail Step 1',
                color: colors.green
            }, {
                name: 'Goal Detail Step 2',
                color: colors.green
            }, {
                name: 'Goal Detail Step 3',
                color: colors.green
            }, {
                name: 'Goal Detail Step 4',
                color: colors.green
            }, {
                name: 'Goal Detail Step 5',
                color: colors.green
            }, {
                name: 'Goal Detail Step 6',
                color: colors.green
            }, {
                name: 'Goal Detail Step 7',
                color: colors.green
            }, {
                name: 'Goal Detail Step 8',
                color: colors.green
            }, {
                name: 'Goal Detail Step 9',
                color: colors.green
            }, {
                name: 'Goal Detail Step 10',
                color: colors.green
            }, {
                name: 'Goal Detail ComparisonType',
                color: colors.purple
            }, {
                name: 'Goal Detail ComparisonValue',
                color: colors.purple
            }, {
                name: 'Goal Event Condition Type',
                color: colors.yellow
            }, {
                name: 'Goal Event Condition MatchType',
                color: colors.yellow
            }, {
                name: 'Goal Event Condition Expression',
                color: colors.yellow
            }, {
                name: 'Goal Event Condition Type',
                color: colors.yellow
            }, {
                name: 'Goal Event Condition MatchType',
                color: colors.yellow
            }, {
                name: 'Goal Event Condition Expression',
                color: colors.yellow
            }, {
                name: 'Goal Event Condition Type',
                color: colors.yellow
            }, {
                name: 'Goal Event Condition MatchType',
                color: colors.yellow
            }, {
                name: 'Goal Event Condition Expression',
                color: colors.yellow
            }, {
                name: 'Goal Event Condition Type',
                color: colors.yellow
            }, {
                name: 'Goal Event Condition ComparisonType',
                color: colors.yellow
            }, {
                name: 'Goal Event Condition ComparisonValue',
                color: colors.yellow
            }];

            break;

        default:
            cells = [];

        }

        return {
            values: [cells.map(function (cell) { return cell.name; })],
            colors: [cells.map(function (cell) { return cell.color || colors.blue; })]
        };
    },

    build: function () {
        var header = this.getHeaderRow();
        var headerLen = header.values[0].length;
        var headerRow = this.sheet.setRowHeight(1, 35).getRange(1, 1, 1, headerLen);
        var dataRange = this.sheet.getRange(2, 1, this.data.length, headerLen);
        var allData = this.sheet.getRange(2, 1, this.sheet.getMaxRows(), headerLen);

        // add header row
        headerRow
            .setBackgrounds(header.colors)
            .setFontColor('white')
            .setFontSize(12)
            .setFontWeight('bold')
            .setVerticalAlignment('middle')
            .setValues(header.values);

        // clear existing data
        if (!dataRange.isBlank()) {
            allData.clearContent();
        }

        // add data to sheet
        dataRange.setValues(this.data);

        // auto resize all columns
        header.values[0].forEach(function (e, i) {
            this.sheet.autoResizeColumn(i + 1);
        }, this);

        // freeze the header row
        this.sheet.setFrozenRows(1);
    }
};

/*
 *
 * API
 *
 */

var Api = function (selectedAccounts) {
    this.selectedAccounts = selectedAccounts;
};

Api.prototype = {
    createFilterRow: function (type, details) {
        var len = 16;
        var arr = Array.call(null, len);

        switch (type) {

        case 'EXCLUDE_OR_INCLUDE':
            arr[0] = details.field;
            arr[1] = details.matchType;
            arr[2] = details.expressionValue;
            arr[3] = details.caseSensitive;
            break;

        case 'UPPERCASE_OR_LOWERCASE':
            arr[0] = details.field;
            break;

        case 'SEARCH_AND_REPLACE':
            arr[0] = details.field;
            arr[3] = details.searchString;
            arr[4] = details.replaceString;
            arr[5] = details.caseSensitive;
            break;

        case 'ADVANCED':
            arr[0] = details.field;
            arr[6] = details.fieldA;
            arr[7] = details.extractA;
            arr[8] = details.fieldB;
            arr[9] = details.extractB;
            arr[10] = details.outputToField;
            arr[11] = details.outputConstructor;
            arr[12] = details.fieldARequired;
            arr[13] = details.fieldBRequired;
            arr[14] = details.overrideOutputField;
            arr[15] = details.caseSensitive;
            break;

        }

        return replaceSparseInArrayWithDefault(arr, '-', len);
    },

    createGoalRow: function (type, details) {
        var len = 28;
        var arr = Array.call(null, len);
        var goalDetailStep = [];
        var conditionDetail = [];
        var steps, conditions;

        switch (type) {

        case 'urlDestinationDetails':
            steps = details.steps;

            if (steps) {
                steps.forEach(function (step) {
                    goalDetailStep.push(step.url);
                });
            }

            arr[0] = details.url;
            arr[1] = details.caseSensitive;
            arr[2] = details.matchType;
            arr[3] = details.firstStepRequired;
            arr[4] = goalDetailStep[0];
            arr[5] = goalDetailStep[1];
            arr[6] = goalDetailStep[2];
            arr[7] = goalDetailStep[3];
            arr[8] = goalDetailStep[4];
            arr[9] = goalDetailStep[5];
            arr[10] = goalDetailStep[6];
            arr[11] = goalDetailStep[7];
            arr[12] = goalDetailStep[8];
            arr[13] = goalDetailStep[9];

            goalDetailStep = [];

            break;

        case 'visitTimeOnSiteDetails_OR_visitNumPagesDetails':
            arr[14] = details.comparisonType;
            arr[15] = details.comparisonValue;

            break;

        case 'eventDetails':
            conditions = details.eventConditions;

            if (conditions) {
                conditions.forEach(function (condition) {
                    if (condition.type === 'VALUE') {
                        conditionDetail.push(
                            condition.type,
                            condition.comparisonType,
                            condition.comparisonValue
                        );
                    } else {
                        conditionDetail.push(
                            condition.type,
                            condition.matchType,
                            condition.expression
                        );
                    }
                });
            }

            arr[16] = conditionDetail[0];
            arr[17] = conditionDetail[1];
            arr[18] = conditionDetail[2];
            arr[19] = conditionDetail[3];
            arr[20] = conditionDetail[4];
            arr[21] = conditionDetail[5];
            arr[22] = conditionDetail[6];
            arr[23] = conditionDetail[7];
            arr[24] = conditionDetail[8];
            arr[25] = conditionDetail[9];
            arr[26] = conditionDetail[10];
            arr[27] = conditionDetail[11];

            conditionDetail = [];

            break;
        }

        return replaceSparseInArrayWithDefault(arr, '-', len);
    },

    wrapperGetViewFilterData: function (account, property, profile, cb) {
        var profileFilters = Analytics.Management.ProfileFilterLinks.list(account, property, profile).getItems();

        validateCallback(cb);

        return cb.call(this, profileFilters);
    },

    getViewFilterData: function (cb) {
        var results = [];

        this.selectedAccounts.forEach(function (account) {
            account.webProperties.forEach(function (property) {
                property.profiles.forEach(function (profile) {
                    this.wrapperGetViewFilterData(account.id, property.id, profile.id, function (filtersList) {
                        filtersList.forEach(function (filter) {
                            results.push([
                                account.name,
                                account.id,
                                property.name,
                                property.id,
                                profile.name,
                                profile.id,
                                filter.filterRef.name,
                                filter.filterRef.id
                            ]);
                        });
                    });
                }, this);
            }, this);
        }, this);

        cb(results);
    },

    wrapperGetViewData: function (account, property, cb) {
        var viewsList = Analytics.Management.Profiles.list(account, property).getItems();

        validateCallback(cb);

        return cb.call(this, viewsList);
    },

    getViewData: function (cb) {
        var results = [];

        this.selectedAccounts.forEach(function (account) {
            account.webProperties.forEach(function (property) {
                this.wrapperGetViewData(account.id, property.id, function (profilesList) {
                    profilesList.forEach(function (profile) {
                        results.push([
                            account.name,
                            account.id,
                            property.name,
                            property.id,
                            profile.name,
                            profile.id,
                            profile.websiteUrl,
                            profile.timezone,
                            profile.defaultPage,
                            profile.excludeQueryParameters,
                            profile.currency,
                            profile.siteSearchQueryParameters,
                            profile.stripSiteSearchQueryParameters,
                            profile.siteSearchCategoryParameters,
                            profile.stripSiteSearchCategoryParameters
                        ]);
                    });
                });
            }, this);
        }, this);

        cb(results);
    },

    wrapperGetAccountFilterData: function (account, cb) {
        var filtersList = Analytics.Management.Filters.list(account).getItems();

        validateCallback(cb);

        return cb.call(this, filtersList);
    },

    getAccountFilterData: function (cb) {
        var details, rowDefaults, rowDetails;
        var results = [];

        this.selectedAccounts.forEach(function (account) {
            this.wrapperGetAccountFilterData(account.id, function (filtersList) {
                filtersList.forEach(function (filter) {
                    rowDefaults = [
                        account.name,
                        account.id,
                        filter.name,
                        filter.id,
                        filter.type
                    ];

                    if (filter.type === 'EXCLUDE' || filter.type === 'INCLUDE') {
                        details = filter.getIncludeDetails() || filter.getExcludeDetails();
                        rowDetails = this.createFilterRow('EXCLUDE_OR_INCLUDE', details);

                        results.push(rowDefaults.concat(rowDetails));
                    }

                    if (filter.type === 'UPPERCASE' || filter.type === 'LOWERCASE') {
                        details = filter.uppercaseDetails || filter.lowercaseDetails;
                        rowDetails = this.createFilterRow('UPPERCASE_OR_LOWERCASE', details);

                        results.push(rowDefaults.concat(rowDetails));
                    }

                    if (filter.type === 'SEARCH_AND_REPLACE') {
                        details = filter.searchAndReplaceDetails;
                        rowDetails = this.createFilterRow('SEARCH_AND_REPLACE', details);

                        results.push(rowDefaults.concat(rowDetails));
                    }

                    if (filter.type === 'ADVANCED') {
                        details = filter.advancedDetails;
                        rowDetails = this.createFilterRow('ADVANCED', details);

                        results.push(rowDefaults.concat(rowDetails));
                    }
                }, this);
            });
        }, this);

        cb(results);
    },

    wrapperGetGoalData: function (account, property, profile, cb) {
        var goalsList = Analytics.Management.Goals.list(account, property, profile).getItems();

        validateCallback(cb);

        return cb.call(this, goalsList);
    },

    getGoalData: function (cb) {
        var details, rowDetails;
        var results = [];

        this.selectedAccounts.forEach(function (account) {
            account.webProperties.forEach(function (property) {
                property.profiles.forEach(function (profile) {
                    this.wrapperGetGoalData(account.id, property.id, profile.id, function (goalsList) {
                        goalsList.forEach(function (goal) {
                            var rowDefaults = [
                                account.name,
                                account.id,
                                property.name,
                                property.id,
                                profile.name,
                                profile.id,
                                goal.name,
                                goal.id,
                                goal.type,
                                goal.active,
                                goal.value
                            ];
                            if (goal.urlDestinationDetails) {
                                details = goal.urlDestinationDetails;
                                rowDetails = this.createGoalRow('urlDestinationDetails', details);

                                results.push(rowDefaults.concat(rowDetails));
                            }

                            if (goal.visitTimeOnSiteDetails || goal.visitNumPagesDetails) {
                                details = goal.visitTimeOnSiteDetails || goal.visitNumPagesDetails;
                                rowDetails = this.createGoalRow('visitTimeOnSiteDetails_OR_visitNumPagesDetails', details);

                                results.push(rowDefaults.concat(rowDetails));
                            }

                            if (goal.eventDetails) {
                                details = goal.eventDetails;
                                rowDetails = this.createGoalRow('eventDetails', details);

                                results.push(rowDefaults.concat(rowDetails));
                            }

                        }, this);
                    });
                }, this);
            }, this);
        }, this);
        cb(results);
    }
};

/*
 *
 * MAIN
 *
 */

function onOpen(e) {
    return SpreadsheetApp
        .getUi()
        .createAddonMenu()
        .addItem('Get Accounts', 'showSidebar')
        .addToUi();
}

function showSidebar() {
    var ui = HtmlService
        .createTemplateFromFile('index')
        .evaluate()
        .setTitle('Auditor')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);

    return SpreadsheetApp.getUi().showSidebar(ui);
}

function getAccountSummaries() {
    var items = Analytics.Management.AccountSummaries.list().getItems();

    if (!items) {
        return [];
    }

    return items;
}

function getReports() {

    return [{
        'name': 'View Filters',
        'id': 1234
    }, {
        'name': 'View Settings',
        'id': 2345
    }, {
        'name': 'Account Filters',
        'id': 3456
    }, {
        'name': 'View Goals',
        'id': 4567
    }];
}

function buildSheetWithData(sheetName, data) {
    var sheet = new Sheet(sheetName, data);

    return sheet.build();
}

function doAuditOfAccounts(selectedAccounts) {
    var api = new Api(selectedAccounts);

    api.getViewFilterData(function (results) {
        buildSheetWithData('Filters - View Level', results);
    });

    api.getViewData(function (results) {
        buildSheetWithData('Views', results);
    });

    api.getAccountFilterData(function (results) {
        buildSheetWithData('Filters - Account Level', results);
    });

    api.getGoalData(function (results) {
        buildSheetWithData('Goals', results);
    });

}

function prepareIncomingAccounts(selectedAccounts) {
    var accountSummaries = getAccountSummaries();
    var selectedIds = selectedAccounts.map(function (account) { return account.value; });
    var accountsObj = accountSummaries.filter(function (account) { return selectedIds.indexOf(account.id) > -1; });

    return doAuditOfAccounts(accountsObj);
}

function jsonAccountSummary() {
    var accountSummaries = getAccountSummaries();

    return JSON.stringify(accountSummaries, ['name', 'id', 'webProperties', 'profiles']);

}