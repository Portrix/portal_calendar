"use strict";
var Portal;
(function (Portal) {
    var Services;
    (function (Services) {
        var Absence = /** @class */ (function () {
            //Календарь руководителей
            function Absence() {
                this.managment = ko.observableArray();
                this.selectedId = ko.observable();
                this.newFormUrl = _spPageContextInfo.webAbsoluteUrl + "/Lists/Absence/NewForm.aspx";
                this.showCalendar = ko.observable(false);
                this.showNone = ko.observable('showNone');
                this.showNoneButton = ko.observable('showNoneButton');
                this.selectedName = ko.observable();
                this.isAdmin = ko.observable();
                $("#loader").addClass("active");
                this.events = [];
                this.selectedId(0);
                this.init = this.init.bind(this);
                this.openNewFormDialog = this.openNewFormDialog.bind(this);
                this.initUI = this.initUI.bind(this);
                this.getManagers = this.getManagers.bind(this);
                this.getManagersSuccessCallback = this.getManagersSuccessCallback.bind(this);
                this.getEvents = this.getEvents.bind(this);
                this.getEventsSuccessCallback = this.getEventsSuccessCallback.bind(this);
                this.deleteListItem = this.deleteListItem.bind(this);
                this.errorCallback = this.errorCallback.bind(this);
                this.showAll = this.showAll.bind(this);
                this.openDispForm = this.openDispForm.bind(this);
                this.filterCalendarView = this.filterCalendarView.bind(this);
                this.openEditForm = this.openEditForm.bind(this);
                this.CheckMemberInAdminGroup = this.CheckMemberInAdminGroup.bind(this);
                this.success = this.success.bind(this);
                this.failure = this.failure.bind(this);
                SP.SOD.executeFunc('sp.js', 'SP.ClientContext', this.init);
            }
            Absence.prototype.CheckMemberInAdminGroup = function () {
                var clientContext = new SP.ClientContext.get_current();
                this.currentUser = clientContext.get_web().get_currentUser();
                clientContext.load(this.currentUser);
                this.userGroups = this.currentUser.get_groups();
                clientContext.load(this.userGroups);
                clientContext.executeQueryAsync(this.success, this.failure);
            };
            Absence.prototype.success = function () {
                var groupsEnumerator = this.userGroups.getEnumerator();
                var isAdmin = false;
                while (groupsEnumerator.moveNext() && !isAdmin) {
                    var group = groupsEnumerator.get_current();
                    if (group.get_title() == "Календарь руководителей") {
                        isAdmin = true;
                    }
                }
                if (isAdmin) {
                    $.model.isAdmin(true);
                }
                else {
                    $.model.isAdmin(false);
                }
            };
            Absence.prototype.failure = function (sender, args) {
                alert('Request failed. ' + args.get_message() +
                    '\n' + args.get_stackTrace());
            };
            Absence.prototype.openNewFormDialog = function (id) {
                var options = new SP.UI.DialogOptions();
                options.title = "Создать:";
                options.url = this.newFormUrl;
                options.allowMaximize = false;
                options.width = 550;
                options.height = 400;
                options.dialogReturnValueCallback = function (result, target) {
                    if (result == 1) {
                        var clientContext = SP.ClientContext.get_current();
                        var oList = clientContext.get_web().get_lists().getByTitle('Absence');
                        var camlQuery = new SP.CamlQuery();
                        camlQuery.set_viewXml("<View><Query><Where>"
                            + "</Where>"
                            + "<OrderBy><FieldRef Name='ID' Ascending='False' /></OrderBy>"
                            + "</Query>"
                            + "<RowLimit>1</RowLimit>"
                            + "</View>");
                        var items = oList.getItems(camlQuery);
                        clientContext.load(items);
                        clientContext.executeQueryAsync(function () {
                            var count = items.get_count();
                            //should only be 1
                            if (count > 1) {
                                throw "Something is wrong. Should only be one latest list item / doc";
                            }
                            var enumerator = items.getEnumerator();
                            enumerator.moveNext();
                            var item = enumerator.get_current();
                            var ev = new Event(item);
                            console.log(ev);
                            $('#calendar').fullCalendar('renderEvent', ev);
                        }, function () {
                            //failure handling comes here
                            alert("failed");
                        });
                    }
                };
                SP.UI.ModalDialog.showModalDialog(options);
            };
            Absence.prototype.init = function () {
                this.getManagers();
                this.getEvents();
                this.CheckMemberInAdminGroup();
            };
            Absence.prototype.showAll = function () {
                this.selectedId(0);
                $(".newCard").removeClass('newColor');
                $("#calendar").fullCalendar("rerenderEvents");
                $.model.showNone('showTrue');
                $.model.selectedName('всех руководителей.');
            };
            Absence.prototype.initUI = function () {
                $('.ui.dropdown').dropdown();
            };
            Absence.prototype.openDispForm = function (id) {
                var pageUrl = _spPageContextInfo.webAbsoluteUrl + "/Lists/Absence/DispForm.aspx?ID=" + id;
                var options = new SP.UI.DialogOptions();
                options.title = "Просмотр формы:";
                options.url = pageUrl;
                options.allowMaximize = false;
                options.width = 550;
                options.height = 400;
                SP.UI.ModalDialog.showModalDialog(options);
            };
            Absence.prototype.openEditForm = function (id) {
                var pageUrl = _spPageContextInfo.webAbsoluteUrl + "/Lists/Absence/EditForm.aspx?ID=" + id;
                var options = new SP.UI.DialogOptions();
                options.title = "Редактирование формы:";
                options.url = pageUrl;
                options.allowMaximize = false;
                options.width = 550;
                options.height = 400;
                options.args = { eventId: id };
                options.dialogReturnValueCallback = function (result, target) {
                    if (result == 1) {
                        var eventId = this.get_args().eventId;
                        var clientContext = new SP.ClientContext();
                        var list = clientContext.get_web().get_lists().getByTitle("Absence");
                        var item = list.getItemById(eventId);
                        clientContext.load(item);
                        clientContext.executeQueryAsync(function () {
                            var ev = new Event(item);
                            var updatedEvent = ev.syncronizeWithCalendar();
                            $('#calendar').fullCalendar('updateEvents', updatedEvent);
                        }, function (sender, args) { alert(args.get_message()); });
                    }
                };
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
                        basicWeek: { buttonText: "Неделя" },
                    },
                    defaultView: 'basicWeek',
                    locale: "ru",
                    displayEventTime: false,
                    events: this.events,
                    eventMouseover: function (event, element, view) {
                        var a = element.currentTarget;
                        var onclickFuncEdit = "$.model.openEditForm(" + event.id + ")";
                        var onclickFuncView = "$.model.openDispForm(" + event.id + ")";
                        var deleteListItem = "$.model.deleteListItem(" + event.id + ")";
                        var icoUrl = _spPageContextInfo.webAbsoluteUrl;
                        var viewButtonHtml = '<span class="editLink" style="cursor:pointer;" onclick="' + onclickFuncView +
                            '"><img class="iconPic" src="' + icoUrl + '/SiteAssets/absence/view.svg" width="20px" height="20px" /></span>';
                        if (_this.isAdmin()) {
                            viewButtonHtml += '<span class="editLink" style="cursor:pointer;" onclick="' +
                                onclickFuncEdit + '"><img class="iconPic" src="' + icoUrl + '/SiteAssets/absence/pencil.svg" width="20px" height="20px"/></span>' +
                                '<span class="deleteLink" style="cursor:pointer;" onclick="' + deleteListItem + '"><img class="iconPic" src="' + icoUrl + '/SiteAssets/absence/delete.svg" width="20px" height="20px"/></span>';
                        }
                        a.insertAdjacentHTML('beforeend', viewButtonHtml);
                    },
                    eventMouseout: function (event, element, view) {
                        $(".editLink").remove();
                        $(".deleteLink").remove();
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
            Absence.prototype.deleteListItem = function (id) {
                var deleteResult = confirm("Вы уверены, что хотите удалить запись?");
                if (deleteResult) {
                    var clientContext = new SP.ClientContext();
                    var list = clientContext.get_web().get_lists().getByTitle("Absence");
                    var item = list.getItemById(id);
                    item.deleteObject();
                    $('#calendar').fullCalendar('removeEvents', id);
                    clientContext.executeQueryAsync(function () {
                        alert("Запись успешно удалена!");
                    }, function (sender, args) { alert(args.get_message()); });
                }
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
                $.model.showNone('showTrue');
                $.model.selectedId(this.userId);
                $("#calendar").fullCalendar("rerenderEvents");
                $.model.selectedName(this.name);
                $(".newCard").removeClass('newColor');
                $('#newCard_' + this.userId).addClass('newColor');
            };
            return Manager;
        }());
        Services.Manager = Manager;
        var Event = /** @class */ (function () {
            function Event(oListItem) {
                this.syncronizeWithCalendar = this.syncronizeWithCalendar.bind(this);
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
            Event.prototype.syncronizeWithCalendar = function () {
                var calEvent = $('#calendar').fullCalendar('clientEvents', this.id)[0];
                calEvent.userId = this.userId;
                calEvent.title = this.title;
                calEvent.status = this.status;
                calEvent.color = this.color;
                calEvent.start = moment(this.start);
                calEvent.end = $.fullCalendar.moment(this.end).add(10, 'seconds');
                return calEvent;
            };
            return Event;
        }());
        Services.Event = Event;
    })(Services = Portal.Services || (Portal.Services = {}));
})(Portal || (Portal = {}));
