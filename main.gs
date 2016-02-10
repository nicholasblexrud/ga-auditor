/*global Logger, SpreadsheetApp, HtmlService, Analytics */

/*
 *
 * Utility/helper functions
 *
 */

var colors = {
    blue: '#4d90fe',
    green: '#0fc357',
    purple: '#c27ba0',
    yellow: '#e7fe2b',
    grey: '#666666',
    white: '#ffffff'
};

function headerValuesAndColors(array) {
    return {
        values: [array.map(function (element) { return element.name; })],
        colors: [array.map(function (element) { return element.color || colors.blue; })]
    };
}

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

var sheet = {
    init: function (config) {
        this.workbook = SpreadsheetApp.getActiveSpreadsheet();
        this.name = config.name;
        this.header = config.header;
        this.headerLength = this.header.values[0].length;
        this.data = config.data;
        this.sheet = this.workbook.getSheetByName(this.name) || this.workbook.insertSheet(this.name);

        return this;
    },

    buildTitle: function (cb) {
        var rowHeight = 35;
        var titleRow = this.sheet.setRowHeight(1, rowHeight).getRange(1, 1, 1, this.headerLength);
        var titleCell = this.sheet.getRange(1, 1);

        titleRow
            .setBackground(colors.grey)
            .setFontColor(colors.white)
            .setFontSize(12)
            .setFontWeight('bold')
            .setVerticalAlignment('middle')
            .setHorizontalAlignment('left')
            .mergeAcross();

        titleCell
            .setValue(this.name);

        cb.call(this);
    },

    buildHeader: function (cb) {
        var rowHeight = 35;
        var headerRow = this.sheet.setRowHeight(2, rowHeight).getRange(2, 1, 1, this.headerLength);

        // add style header row
        headerRow
            .setBackgrounds(this.header.colors)
            .setFontColor(colors.white)
            .setFontSize(12)
            .setFontWeight('bold')
            .setVerticalAlignment('middle')
            .setValues(this.header.values);

        // freeze the header row
        this.sheet.setFrozenRows(2);

        cb.call(this);

    },

    buildData: function (cb) {
        var dataRange = this.sheet.getRange(3, 1, this.data.length, this.headerLength);
        var allData = this.sheet.getRange(3, 1, this.sheet.getMaxRows(), this.headerLength);

        // clear existing data
        if (!dataRange.isBlank()) {
            allData.clearContent();
        }

        // add data to sheet
        dataRange.setValues(this.data);

        cb.call(this);
    },

    cleanup: function () {
        // auto resize all columns
        this.header.values[0].forEach(function (e, i) {
            this.sheet.autoResizeColumn(i + 1);
        }, this);
    },

    build: function () {
        this.buildTitle(function () {
            this.buildHeader(function () {
                this.buildData(function () {
                    this.cleanup();
                });
            });
        });
    }
};

/*
 *
 * API
 *
 */

var api = {

    goals: {
        init: function (config) {
            this.account = config.account;
            this.accountName = this.account.name;

            this.header = this.getHeader();

            return this;
        },
        name: 'Goals',
        getHeader: function () {
            var data = [{
                name: 'Property Name'
            }, {
                name: 'Property Id'
            }, {
                name: 'View Name'
            }, {
                name: 'Goal Name'
            }, {
                name: 'Goal Id'
            }, {
                name: 'Type'
            }, {
                name: 'Active?'
            }, {
                name: 'Value'
            }, {
                name: 'Goal URL',
                color: colors.green
            }, {
                name: 'Goal CaseSensitive',
                color: colors.green
            }, {
                name: 'Match Type',
                color: colors.green
            }, {
                name: 'Required Step?',
                color: colors.green
            }, {
                name: 'Step 1',
                color: colors.green
            }, {
                name: 'Step 2',
                color: colors.green
            }, {
                name: 'Step 3',
                color: colors.green
            }, {
                name: 'Step 4',
                color: colors.green
            }, {
                name: 'Step 5',
                color: colors.green
            }, {
                name: 'Step 6',
                color: colors.green
            }, {
                name: 'Step 7',
                color: colors.green
            }, {
                name: 'Step 8',
                color: colors.green
            }, {
                name: 'Step 9',
                color: colors.green
            }, {
                name: 'Step 10',
                color: colors.green
            }, {
                name: 'ComparisonType',
                color: colors.purple
            }, {
                name: 'ComparisonValue',
                color: colors.purple
            }, {
                name: 'Event Category type',
                color: colors.yellow
            }, {
                name: 'Event Category condition',
                color: colors.yellow
            }, {
                name: 'Event Category value',
                color: colors.yellow
            }, {
                name: 'Event Action type',
                color: colors.yellow
            }, {
                name: 'Event Action condition',
                color: colors.yellow
            }, {
                name: 'Event Action value',
                color: colors.yellow
            }, {
                name: 'Event Label type',
                color: colors.yellow
            }, {
                name: 'Event Label condition',
                color: colors.yellow
            }, {
                name: 'Event Label value',
                color: colors.yellow
            }, {
                name: 'Event Value type',
                color: colors.yellow
            }, {
                name: 'Event Value condition',
                color: colors.yellow
            }, {
                name: 'Event Value value',
                color: colors.yellow
            }];

            return headerValuesAndColors(data);
        },
        row: function (type, details) {
            var len = 26;
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
        wrapper: function (account, property, profile, cb) {
            var goalsList = Analytics.Management.Goals.list(account, property, profile).getItems();

            validateCallback(cb);

            return cb.call(this, goalsList);
        },
        getData: function (cb) {
            var details, rowDetails;
            var results = [];

            this.account.webProperties.forEach(function (property) {
                property.profiles.forEach(function (profile) {
                    this.wrapper(this.account.id, property.id, profile.id, function (goalsList) {
                        goalsList.forEach(function (goal) {
                            var rowDefaults = [
                                property.name,
                                property.id,
                                profile.name,
                                goal.name,
                                goal.id,
                                goal.type,
                                goal.active,
                                goal.value
                            ];
                            if (goal.urlDestinationDetails) {
                                details = goal.urlDestinationDetails;
                                rowDetails = this.row('urlDestinationDetails', details);

                                results.push(rowDefaults.concat(rowDetails));
                            }

                            if (goal.visitTimeOnSiteDetails || goal.visitNumPagesDetails) {
                                details = goal.visitTimeOnSiteDetails || goal.visitNumPagesDetails;
                                rowDetails = this.row('visitTimeOnSiteDetails_OR_visitNumPagesDetails', details);

                                results.push(rowDefaults.concat(rowDetails));
                            }

                            if (goal.eventDetails) {
                                details = goal.eventDetails;
                                rowDetails = this.row('eventDetails', details);

                                results.push(rowDefaults.concat(rowDetails));
                            }

                        }, this);
                    });
                }, this);
            }, this);
            cb(results);
        }
    },

    settings: {
        init: function (config) {
            this.account = config.account;
            this.accountName = this.account.name;

            this.header = this.getHeader();

            return this;
        },
        name: 'Settings',
        getHeader: function () {
            var data = [{
                name: 'Property Name'
            }, {
                name: 'Property Id'
            }, {
                name: 'View Name'
            }, {
                name: 'Website URL'
            }, {
                name: 'Time zone '
            }, {
                name: 'Default Page'
            }, {
                name: 'Exclude Query Parameters'
            }, {
                name: 'Currency'
            }, {
                name: 'Site Search Query Parameters'
            }, {
                name: 'Strip Site Search Query Parameters'
            }, {
                name: 'Site Search Category Parameters',
            }, {
                name: 'Strip Site Search Category Parameters'
            }];

            return headerValuesAndColors(data);
        },
        wrapper: function (account, property, cb) {
            var viewsList = Analytics.Management.Profiles.list(account, property).getItems();

            validateCallback(cb);

            return cb.call(this, viewsList);
        },
        getData: function (cb) {
            var results = [];

            this.account.webProperties.forEach(function (property) {
                this.wrapper(this.account.id, property.id, function (profilesList) {
                    profilesList.forEach(function (profile) {
                        results.push([
                            property.name,
                            property.id,
                            profile.name,
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
                    }, this);
                });
            }, this);

            cb(results);
        }
    },

    filters: {
        init: function (config) {
            this.account = config.account;
            this.accountName = this.account.name;

            this.header = this.getHeader();

            return this;
        },
        name: 'Filters',
        getHeader: function () {
            var data = [{
                name: 'Property Name'
            }, {
                name: 'Property Id'
            }, {
                name: 'View Name'
            },  {
                name: 'Filter Name'
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

            return headerValuesAndColors(data);
        },
        row: function (type, details) {
            var len = 14;
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

        getData: function (cb) {
            this.links(function (links) {
                this.lists(function (lists) {

                    //TODO: adjust column header
                    links.forEach(function (link) {
                        var linkFilterId = link[4];

                        lists.forEach(function (list) {
                            var listFilterId = list[0];
                            if (linkFilterId === listFilterId) {
                                link[4] = list[1];
                                link[5] = list[2];
                                link[6] = list[3];
                                link[7] = list[4];
                                link[8] = list[5];
                                link[9] = list[6];
                                link[10] = list[7];
                                link[11] = list[8];
                                link[12] = list[9];
                                link[13] = list[10];
                                link[14] = list[11];
                                link[15] = list[12];
                                link[16] = list[13];
                                link[17] = list[14];
                                link[18] = list[15];
                                link[19] = list[16];
                                link[20] = list[17];
                            }
                        });
                    });
                });

                cb(links);
            });
        },

        wrapperLinks: function (account, property, profile, cb) {
            var links = Analytics.Management.ProfileFilterLinks.list(account, property, profile).getItems();

            validateCallback(cb);

            return cb.call(this, links);
        },

        wrapperLists: function (account, cb) {
            var list = Analytics.Management.Filters.list(account).getItems();

            validateCallback(cb);

            return cb.call(this, list);
        },

        links: function (cb) {
            var results = [];

            this.account.webProperties.forEach(function (property) {
                property.profiles.forEach(function (profile) {
                    this.wrapperLinks(this.account.id, property.id, profile.id, function (links) {
                        links.forEach(function (filter) {
                            results.push([
                                property.name,
                                property.id,
                                profile.name,
                                filter.filterRef.name,
                                filter.filterRef.id
                            ]);
                        }, this);
                    });
                }, this);
            }, this);

            return cb.call(this, results);
        },

        lists: function (cb) {
            var details, rowDefaults, rowDetails;
            var results = [];

            this.wrapperLists(this.account.id, function (list) {
                list.forEach(function (filter) {
                    rowDefaults = [
                        filter.id,
                        filter.type
                    ];

                    if (filter.type === 'EXCLUDE' || filter.type === 'INCLUDE') {
                        details = filter.getIncludeDetails() || filter.getExcludeDetails();
                        rowDetails = this.row('EXCLUDE_OR_INCLUDE', details);

                        results.push(rowDefaults.concat(rowDetails));
                    }

                    if (filter.type === 'UPPERCASE' || filter.type === 'LOWERCASE') {
                        details = filter.uppercaseDetails || filter.lowercaseDetails;
                        rowDetails = this.row('UPPERCASE_OR_LOWERCASE', details);

                        results.push(rowDefaults.concat(rowDetails));
                    }

                    if (filter.type === 'SEARCH_AND_REPLACE') {
                        details = filter.searchAndReplaceDetails;
                        rowDetails = this.row('SEARCH_AND_REPLACE', details);

                        results.push(rowDefaults.concat(rowDetails));
                    }

                    if (filter.type === 'ADVANCED') {
                        details = filter.advancedDetails;
                        rowDetails = this.row('ADVANCED', details);

                        results.push(rowDefaults.concat(rowDetails));
                    }
                }, this);
            });

            return cb.call(this, results);
        }
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

function getReports() {

    return JSON.stringify([{
        'name': 'Filters',
        'id': 'filters'
    }, {
        'name': 'Settings',
        'id': 'settings'
    }, {
        'name': 'Goals',
        'id': 'goals'
    }]);
}

function generateReport(account, reportName) {
    var report = api[reportName];

    report.init({ account: account })
        .getData(function (data) {
            sheet.init({
                'name': account.name + ': ' + report.name,
                'header': report.header,
                'data': data
            })
                .build();
        });
}

function saveReportDataFromSidebar(data) {
    var parsed = JSON.parse(data);

    return generateReport(parsed.ids, parsed.report);
}

function getAccountSummary() {
    var items = Analytics.Management.AccountSummaries.list().getItems();

    if (!items) {
        return [];
    }

    return JSON.stringify(items, ['name', 'id', 'webProperties', 'profiles']);
}