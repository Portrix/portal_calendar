module Portal.Services {
    declare var $: any;
    declare var SP: any;
    declare var Function: any;
    declare var _spPageContextInfo: any;
    declare var ko: any;
    declare var moment: any;
    declare var ExecuteOrDelayUntilScriptLoaded: any;

    export class Absence {
        public managment = ko.observableArray();
        public events: Array<Event>;
        public eventsListItems: any;
        public managersListItems: any;
        public selectedId = ko.observable();
        public newFormUrl = _spPageContextInfo.webAbsoluteUrl + "/Lists/Absence/NewForm.aspx"
        public showCalendar = ko.observable(false);
        public showNone = ko.observable('showNone');
        public showNoneButton = ko.observable('showNoneButton');
        public selectedName = ko.observable();
        public currentUser: any;
        public userGroups: any;
        public isAdmin = ko.observable();
        //Календарь руководителей
        constructor() {
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
        CheckMemberInAdminGroup() {
            var clientContext = new SP.ClientContext.get_current();
            this.currentUser = clientContext.get_web().get_currentUser();
            clientContext.load(this.currentUser);

            this.userGroups = this.currentUser.get_groups();
            clientContext.load(this.userGroups);
            clientContext.executeQueryAsync(this.success, this.failure);
        }
        success() {
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
            } else {
                $.model.isAdmin(false);

            }

        }
        failure(sender: any, args: any) {
            alert('Request failed. ' + args.get_message() +
                '\n' + args.get_stackTrace());



        }

        openNewFormDialog(id: any) {

            var options = new SP.UI.DialogOptions();
            options.title = "Создать:";
            options.url = this.newFormUrl;
            options.allowMaximize = false;
            options.width = 550;
            options.height = 400;
            options.dialogReturnValueCallback = function (result: any, target: any) {
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
            }
            SP.UI.ModalDialog.showModalDialog(options);
        }

        init() {
            this.getManagers();
            this.getEvents();
            this.CheckMemberInAdminGroup();
        }
        showAll() {
            this.selectedId(0);
            $(".newCard").removeClass('newColor');
            $("#calendar").fullCalendar("rerenderEvents");
            $.model.showNone('showTrue');
            $.model.selectedName('всех руководителей.');

        }
        initUI() {
            $('.ui.dropdown').dropdown();
        }
        openDispForm(id: any) {
            var pageUrl = _spPageContextInfo.webAbsoluteUrl + "/Lists/Absence/DispForm.aspx?ID=" + id;
            var options = new SP.UI.DialogOptions();
            options.title = "Просмотр формы:";
            options.url = pageUrl;
            options.allowMaximize = false;
            options.width = 550;
            options.height = 400;
            SP.UI.ModalDialog.showModalDialog(options);
        }
        openEditForm(id: any) {
            var pageUrl = _spPageContextInfo.webAbsoluteUrl + "/Lists/Absence/EditForm.aspx?ID=" + id;
            var options = new SP.UI.DialogOptions();
            options.title = "Редактирование формы:";
            options.url = pageUrl;
            options.allowMaximize = false;
            options.width = 550;
            options.height = 400;
            options.args = { eventId: id };
            options.dialogReturnValueCallback = function (result: any, target: any) {
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
                    }, function (sender: any, args: any) { alert(args.get_message()); }
                    );
                }
            };
            SP.UI.ModalDialog.showModalDialog(options);
        }

        filterCalendarView(ev: Event) {
            if (this.selectedId() == 0) {
                return true;
            }
            return ev.userId == this.selectedId();
        }

        getManagers() {
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
        }

        getManagersSuccessCallback() {
            var listItemEnumerator = this.managersListItems.getEnumerator();

            while (listItemEnumerator.moveNext()) {
                var oListItem = listItemEnumerator.get_current();
                this.managment.push(new Manager(oListItem));
                this.managment = this.managment.sort((m1: Manager, m2: Manager) => { return m1.index - m2.index; });
            }
        }

        getEvents() {
            var self = this;
            SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
                var clientContext = SP.ClientContext.get_current();
                var oList = clientContext.get_web().get_lists().getByTitle('Absence');
                var camlQuery = new SP.CamlQuery();
                self.eventsListItems = oList.getItems(camlQuery);
                clientContext.load(self.eventsListItems);
                clientContext.executeQueryAsync(Function.createDelegate(self, self.getEventsSuccessCallback), Function.createDelegate(self, self.errorCallback));
            });
        }

        getEventsSuccessCallback(sender: any, args: any) {
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

                eventMouseover: (event: any, element: any, view: any) => {
                    var a = element.currentTarget;
                    var onclickFuncEdit = "$.model.openEditForm(" + event.id + ")";
                    var onclickFuncView = "$.model.openDispForm(" + event.id + ")";
                    var deleteListItem = "$.model.deleteListItem(" + event.id + ")";
                    var icoUrl = _spPageContextInfo.webAbsoluteUrl;
                    var viewButtonHtml = '<span class="editLink" style="cursor:pointer;" onclick="' + onclickFuncView +
                        '"><img class="iconPic" src="' + icoUrl + '/SiteAssets/absence/view.svg" width="20px" height="20px" /></span>';
                    if (this.isAdmin()) {
                        viewButtonHtml += '<span class="editLink" style="cursor:pointer;" onclick="' +
                            onclickFuncEdit + '"><img class="iconPic" src="' + icoUrl + '/SiteAssets/absence/pencil.svg" width="20px" height="20px"/></span>' +
                            '<span class="deleteLink" style="cursor:pointer;" onclick="' + deleteListItem + '"><img class="iconPic" src="' + icoUrl + '/SiteAssets/absence/delete.svg" width="20px" height="20px"/></span>'
                    }
                    a.insertAdjacentHTML('beforeend', viewButtonHtml);
                },
                eventMouseout: (event: any, element: any, view: any) => {
                    $(".editLink").remove();
                    $(".deleteLink").remove();
                },
                eventRender: (event: any, element: any, view: any) => {
                    return this.filterCalendarView(event);
                },
            });

            $("#loader").removeClass("active");
            $(".kor-semantic").show();
        }

        errorCallback(sender: any, args: any) {
            alert('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        }

        deleteListItem(id: any) {
            var deleteResult = confirm("Вы уверены, что хотите удалить запись?");
            if (deleteResult) {
                var clientContext = new SP.ClientContext();
                var list = clientContext.get_web().get_lists().getByTitle("Absence");
                var item = list.getItemById(id);
                item.deleteObject();
                $('#calendar').fullCalendar('removeEvents', id);
                clientContext.executeQueryAsync(function () {
                    alert("Запись успешно удалена!");
                }, function (sender: any, args: any) { alert(args.get_message()); });
            }
        }
    }

    export class Manager implements IManager {
        userId: number;
        name: string;
        imageUrl: string;
        position: string;
        index: number;
        isSelected = ko.observable();

        constructor(oListItem: any) {
            this.filter = this.filter.bind(this);
            this.userId = oListItem.get_id();
            this.name = oListItem.get_item('Title');
            this.imageUrl = oListItem.get_item('ImageUrl');
            this.isSelected(false);
            this.position = oListItem.get_item('Position');
            this.index = oListItem.get_item('Index');
        }

        filter() {
            $.model.showNone('showTrue');
            $.model.selectedId(this.userId);
            $("#calendar").fullCalendar("rerenderEvents");
            $.model.selectedName(this.name);
            $(".newCard").removeClass('newColor');
            $('#newCard_' + this.userId).addClass('newColor');
        }
    }

    export class Event {
        id: number;
        title: string;
        userId: number;//User
        start: any;//StartDate
        end: any;//EndDate
        status: string;
        color: string;

        constructor(oListItem: any) {
            this.syncronizeWithCalendar = this.syncronizeWithCalendar.bind(this);
            this.id = oListItem.get_id();
            this.userId = oListItem.get_item('ManagerLookup').get_lookupId();
            this.title = oListItem.get_item('ManagerLookup').get_lookupValue();
            this.start = moment(oListItem.get_item('StartDate')).format('YYYY-MM-DD');
            this.end = `${moment(oListItem.get_item('EndDate')).format('YYYY-MM-DD')}T23:59:00`;
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

        syncronizeWithCalendar() {
            var calEvent = $('#calendar').fullCalendar('clientEvents', this.id)[0];
            calEvent.userId = this.userId;
            calEvent.title = this.title;
            calEvent.status = this.status;
            calEvent.color = this.color;
            calEvent.start = moment(this.start);
            calEvent.end = $.fullCalendar.moment(this.end).add(10, 'seconds');
            return calEvent;
        }
    }

    export interface IManager {
        userId: number;
        name: string;
        imageUrl: string;
        position: string;
        index: number;
        isSelected: any;
    }
}