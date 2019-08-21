"use strict";
var Portal;
(function (Portal) {
    var Services;
    (function (Services) {
        var Absence = /** @class */ (function () {
            function Absence() {
                this.managment = ko.observableArray();
                this.selectedId = ko.observable();
                $("#loader").addClass("active");
                this.events = [];
                this.selectedId(0);
                this.init = this.init.bind(this);
                this.initUI = this.initUI.bind(this);
                this.getManagers = this.getManagers.bind(this);
                this.getManagersSuccessCallback = this.getManagersSuccessCallback.bind(this);
                this.getEvents = this.getEvents.bind(this);
                this.getEventsSuccessCallback = this.getEventsSuccessCallback.bind(this);
                this.errorCallback = this.errorCallback.bind(this);
                this.showAll = this.showAll.bind(this);
                this.openDispForm = this.openDispForm.bind(this);
                this.filterCalendarView = this.filterCalendarView.bind(this);
                this.init();
            }
            Absence.prototype.init = function () {
                this.getManagers();
                this.getEvents();
            };
            Absence.prototype.showAll = function () {
                this.selectedId(0);
                $("#calendar").fullCalendar("rerenderEvents");
            };
            Absence.prototype.initUI = function () {
                $('.ui.dropdown').dropdown();
            };
            Absence.prototype.openDispForm = function (ev) {
                var pageUrl = _spPageContextInfo.webAbsoluteUrl + "/Lists/Absence/DispForm.aspx?ID=" + ev.id;
                var options = new SP.UI.DialogOptions();
                options.title = ev.title;
                options.url = pageUrl;
                options.allowMaximize = false;
                options.width = 450;
                options.height = 300;
                SP.UI.ModalDialog.showModalDialog(options);
            };
            Absence.prototype.filterCalendarView = function (ev) {
                if (this.selectedId() == 0) {
                    return true;
                }
                return ev.userId == this.selectedId();
            };
            Absence.prototype.getManagers = function () {
                var self = this;
                SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
                    var clientContext = SP.ClientContext.get_current();
                    var oList = clientContext.get_web().get_lists().getByTitle('Management');
                    var camlQuery = new SP.CamlQuery();
                    camlQuery.set_viewXml("<View><Query><Where><Eq><FieldRef Name='ShowInAbsence'/><Value Type='Boolean'>1</Value></Eq></Where></Query></View>");
                    self.managersListItems = oList.getItems(camlQuery);
                    clientContext.load(self.managersListItems);
                    clientContext.executeQueryAsync(Function.createDelegate(self, self.getManagersSuccessCallback), Function.createDelegate(self, self.errorCallback));
                });
            };
            Absence.prototype.getManagersSuccessCallback = function () {
                var listItemEnumerator = this.managersListItems.getEnumerator();
                while (listItemEnumerator.moveNext()) {
                    var oListItem = listItemEnumerator.get_current();
                    this.managment.push(new Manager(oListItem));
                    this.managment = this.managment.sort(function (m1, m2) { return m1.index - m2.index; });
                }
            };
            Absence.prototype.getEvents = function () {
                var self = this;
                SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
                    var clientContext = SP.ClientContext.get_current();
                    var oList = clientContext.get_web().get_lists().getByTitle('Absence');
                    var camlQuery = new SP.CamlQuery();
                    self.eventsListItems = oList.getItems(camlQuery);
                    clientContext.load(self.eventsListItems);
                    clientContext.executeQueryAsync(Function.createDelegate(self, self.getEventsSuccessCallback), Function.createDelegate(self, self.errorCallback));
                });
            };
            Absence.prototype.getEventsSuccessCallback = function (sender, args) {
                var _this = this;
                var listItemEnumerator = this.eventsListItems.getEnumerator();
                while (listItemEnumerator.moveNext()) {
                    var oListItem = listItemEnumerator.get_current();
                    this.events.push(new Event(oListItem));
                }
                $("#calendar").fullCalendar({
                    header: {
                        left: "prev,next today",
                        center: "title",
                        right: "month,basicWeek"
                    },
                    views: {
                        month: { buttonText: "Месяц" },
                        basicWeek: { buttonText: "Неделя" }
                    },
                    locale: "ru",
                    displayEventTime: false,
                    events: this.events,
                    eventClick: function (event) {
                        _this.openDispForm(event);
                        return false;
                    },
                    eventRender: function (event, element, view) {
                        return _this.filterCalendarView(event);
                    },
                });
                $("#loader").removeClass("active");
                $(".kor-semantic").show();
            };
            Absence.prototype.errorCallback = function (sender, args) {
                alert('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
            };
            return Absence;
        }());
        Services.Absence = Absence;
        var Manager = /** @class */ (function () {
            function Manager(oListItem) {
                this.isSelected = ko.observable();
                this.filter = this.filter.bind(this);
                this.userId = oListItem.get_id();
                this.name = oListItem.get_item('Title');
                this.imageUrl = oListItem.get_item('ImageUrl');
                this.isSelected(false);
                this.position = oListItem.get_item('Position');
                this.index = oListItem.get_item('Index');
            }
            Manager.prototype.filter = function () {
                $.model.selectedId(this.userId);
                $("#calendar").fullCalendar("rerenderEvents");
            };
            return Manager;
        }());
        Services.Manager = Manager;
        var Event = /** @class */ (function () {
            function Event(oListItem) {
                this.id = oListItem.get_id();
                this.userId = oListItem.get_item('ManagerLookup').get_lookupId();
                this.title = oListItem.get_item('ManagerLookup').get_lookupValue();
                this.start = moment(oListItem.get_item('StartDate')).format('YYYY-MM-DD');
                this.end = moment(oListItem.get_item('EndDate')).format('YYYY-MM-DD') + "T23:59:00";
                this.status = oListItem.get_item('Status');
                this.color = "";
                switch (this.status) {
                    case "Не доступен":
                        this.color = "#FEC5BA";
                        break;
                    case "Доступен Спб":
                        this.color = "#D2F2C7";
                        break;
                    case "Доступен Мск":
                        this.color = "#FFF9C0";
                        break;
                }
            }
            return Event;
        }());
        Services.Event = Event;
        var MockManager = /** @class */ (function () {
            function MockManager(id, name, login, imageUrl, position) {
                this.isSelected = ko.observable();
                this.filter = this.filter.bind(this);
                this.userId = id;
                this.name = name;
                this.imageUrl = imageUrl;
                this.isSelected(false);
                this.position = position;
                this.index = 0;
            }
            MockManager.prototype.filter = function () {
                $.model.selectedId(this.userId);
                $("#calendar").fullCalendar("rerenderEvents");
            };
            return MockManager;
        }());
        Services.MockManager = MockManager;
    })(Services = Portal.Services || (Portal.Services = {}));
})(Portal || (Portal = {}));
