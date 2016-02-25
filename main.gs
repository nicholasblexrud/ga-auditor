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
    white: '#ffffff',
    red: '#e06666'
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

function replaceUndefinedInArrayWithDefault(array, defaultValue) {
    return array.map(function (el) {
        return el === undefined ? defaultValue : el;
    });
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
                name: 'Goal Case Sensitive?',
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
                name: 'Event Category condition',
                color: colors.yellow
            }, {
                name: 'Event Category value',
                color: colors.yellow
            }, {
                name: 'Event Action condition',
                color: colors.yellow
            }, {
                name: 'Event Action value',
                color: colors.yellow
            }, {
                name: 'Event Label condition',
                color: colors.yellow
            }, {
                name: 'Event Label value',
                color: colors.yellow
            }, {
                name: 'Event Value condition',
                color: colors.yellow
            }, {
                name: 'Event Value value',
                color: colors.yellow
            }, {
                name: 'Time on Site (in seconds)',
                color: colors.red
            }, {
                name: 'Time on Site value',
                color: colors.red
            }, {
                name: 'Pages per visit',
                color: colors.purple
            }, {
                name: 'Number of pages',
                color: colors.purple
            }];

            return headerValuesAndColors(data);
        },
        row: function (type, details) {
            var len = 26;
            var arr = Array.call(null, len);
            var goalDetailStep = [];
            var steps, conditions;

            switch (type) {

            case 'urlDestinationDetails':
                steps = details.steps;

                if (steps) {
                    steps.forEach(function (step) {
                        goalDetailStep.push(step.url);
                    });
                }

                goalDetailStep = replaceSparseInArrayWithDefault(goalDetailStep, '-', 10);

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

            case 'eventDetails':
                conditions = details.eventConditions;

                if (conditions) {
                    conditions.forEach(function (condition) {

                        if (condition.type === 'CATEGORY') {
                            arr[14] = condition.matchType;
                            arr[15] = condition.expression;

                        } else if (condition.type === 'ACTION') {
                            arr[16] = condition.matchType;
                            arr[17] = condition.expression;

                        } else if (condition.type === 'LABEL') {
                            arr[18] = condition.matchType;
                            arr[19] = condition.expression;

                        } else if (condition.type === 'VALUE') {
                            arr[20] = condition.comparisonType;
                            arr[21] = condition.comparisonValue;

                        }
                    });
                }
                break;

            case 'visitTimeOnSiteDetails':
                arr[22] = details.comparisonType;
                arr[23] = details.comparisonValue;
                break;

            case 'visitNumPagesDetails':
                arr[24] = details.comparisonType;
                arr[25] = details.comparisonValue;
                break;
            }

            return replaceSparseInArrayWithDefault(arr, '-', len);
        },
        wrapper: function (account, property, profile, cb) {
            var goalsList = Analytics.Management.Goals.list(account, property, profile).getItems();
            return cb.call(this, goalsList);
        },
        getData: function (cb) {
            var rowDetails;
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
                                rowDetails = this.row('urlDestinationDetails', goal.urlDestinationDetails);
                            }

                            if (goal.visitTimeOnSiteDetails) {
                                rowDetails = this.row('visitTimeOnSiteDetails', goal.visitTimeOnSiteDetails);
                            }

                            if (goal.visitNumPagesDetails) {
                                rowDetails = this.row('visitNumPagesDetails', goal.visitNumPagesDetails);
                            }

                            if (goal.eventDetails) {
                                rowDetails = this.row('eventDetails', goal.eventDetails);
                            }

                            results.push(rowDefaults.concat(rowDetails));

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
                name: 'Bot Filtering?'
            }, {
                name: 'Site Search Query Parameters',
                color: colors.green
            }, {
                name: 'Strip Site Search Query Parameters',
                color: colors.green
            }, {
                name: 'Site Search Category Parameters',
                color: colors.green
            }, {
                name: 'Strip Site Search Category Parameters',
                color: colors.green
            }, {
                name: 'Ecommerce Tracking?',
                color: colors.yellow
            }, {
                name: 'Enhanced Ecommerce Tracking?',
                color: colors.yellow
            }];

            return headerValuesAndColors(data);
        },
        wrapper: function (account, property, cb) {
            var viewsList = Analytics.Management.Profiles.list(account, property).getItems();
            return cb.call(this, viewsList);
        },
        getData: function (cb) {
            var results = [];

            this.account.webProperties.forEach(function (property) {
                this.wrapper(this.account.id, property.id, function (profilesList) {
                    profilesList.forEach(function (profile) {
                        var rowDefaults = [
                            property.name,
                            property.id,
                            profile.name,
                            profile.websiteUrl,
                            profile.timezone,
                            profile.defaultPage,
                            profile.excludeQueryParameters,
                            profile.currency,
                            profile.botFilteringEnabled,
                            profile.siteSearchQueryParameters,
                            profile.stripSiteSearchQueryParameters,
                            profile.siteSearchCategoryParameters,
                            profile.stripSiteSearchCategoryParameters,
                            profile.eCommerceTracking,
                            profile.enhancedECommerceTracking
                        ];

                        rowDefaults = replaceUndefinedInArrayWithDefault(rowDefaults, '-');

                        results.push(rowDefaults);
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
                name: 'Filter Case Sensitive?'
            }, {
                name: 'Match Type',
                color: colors.green
            }, {
                name: 'Expression Value',
                color: colors.green
            }, {
                name: 'Search Value',
                color: colors.yellow
            }, {
                name: 'Replace Value',
                color: colors.yellow
            }, {
                name: 'Field A',
                color: colors.red
            }, {
                name: 'Extract A Pattern',
                color: colors.red
            }, {
                name: 'Field B',
                color: colors.red
            }, {
                name: 'Extract B Pattern',
                color: colors.red
            }, {
                name: 'Output to Field',
                color: colors.red
            }, {
                name: 'Output to Constructor',
                color: colors.red
            }, {
                name: 'Field A Required?',
                color: colors.red
            }, {
                name: 'Field B Required?',
                color: colors.red
            }, {
                name: 'Override Output Field?',
                color: colors.red
            }];

            return headerValuesAndColors(data);
        },
        row: function (type, details) {
            var len = 15;
            var arr = Array.call(null, len);

            switch (type) {

            case 'EXCLUDE_OR_INCLUDE':
                arr[0] = details.field;
                arr[1] = details.caseSensitive;
                arr[2] = details.matchType;
                arr[3] = details.expressionValue;
                break;

            case 'UPPERCASE_OR_LOWERCASE':
                arr[0] = details.field;
                break;

            case 'SEARCH_AND_REPLACE':
                arr[0] = details.field;
                arr[1] = details.caseSensitive;
                arr[4] = details.searchString;
                arr[5] = details.replaceString;
                break;

            case 'ADVANCED':
                arr[1] = details.caseSensitive;
                arr[6] = details.fieldA;
                arr[7] = details.extractA;
                arr[8] = details.fieldB;
                arr[9] = details.extractB;
                arr[10] = details.outputToField;
                arr[11] = details.outputConstructor;
                arr[12] = details.fieldARequired;
                arr[13] = details.fieldBRequired;
                arr[14] = details.overrideOutputField;
                break;

            }

            return replaceSparseInArrayWithDefault(arr, '-', len);
        },

        getData: function (cb) {
            this.links(function (links) {
                this.lists(function (lists) {

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
                            }
                        });
                    });
                });

                cb(links);
            });
        },

        wrapperLinks: function (account, property, profile, cb) {
            var links = Analytics.Management.ProfileFilterLinks.list(account, property, profile).getItems();
            return cb.call(this, links);
        },

        wrapperLists: function (account, cb) {
            var list = Analytics.Management.Filters.list(account).getItems();
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
                    }

                    if (filter.type === 'UPPERCASE' || filter.type === 'LOWERCASE') {
                        details = filter.uppercaseDetails || filter.lowercaseDetails;
                        rowDetails = this.row('UPPERCASE_OR_LOWERCASE', details);
                    }

                    if (filter.type === 'SEARCH_AND_REPLACE') {
                        rowDetails = this.row('SEARCH_AND_REPLACE', filter.searchAndReplaceDetails);
                    }

                    if (filter.type === 'ADVANCED') {
                        rowDetails = this.row('ADVANCED', filter.advancedDetails);
                    }

                    results.push(rowDefaults.concat(rowDetails));

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
        .addItem('Create report', 'showSidebar')
        .addToUi();
}

function onInstall(e) {
    onOpen(e);
}

function showSidebar() {
    var ui = HtmlService
        .createTemplateFromFile('index')
        .evaluate()
        .setTitle('GA Auditor')
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

    report
        .init({ account: account })
        .getData(function (data) {

            // check if there is any data to be returned
            if (data[0] === undefined || !data[0].length) {
                throw new Error('No data found for ' + reportName + ' in ' + account.name + '. Please try another account or report.');
            }

            sheet.init({
                'name': account.name + ': ' + report.name,
                'header': report.header,
                'data': data
            }).build();
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