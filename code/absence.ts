module Portal.Services {
    declare var $: any;
    declare var SP: any;
    declare var Function: any;
    declare var _spPageContextInfo: any;
    declare var ko: any;
    declare var moment: any;

    export class Absence {
        public managment = ko.observableArray();
        public events: Array<Event>;
        public eventsListItems: any;
        public managersListItems: any;
        public selectedId = ko.observable();

        //test

        constructor() {
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

        init() {
            this.getManagers();
            this.getEvents();
        }

        showAll() {
            this.selectedId(0);
            $("#calendar").fullCalendar("rerenderEvents");
        }

        initUI() {
            $('.ui.dropdown').dropdown();
        }

        openDispForm(ev: Event) {
            var pageUrl = _spPageContextInfo.webAbsoluteUrl + "/Lists/Absence/DispForm.aspx?ID=" + ev.id;

            var options = new SP.UI.DialogOptions();
            options.title = ev.title;
            options.url = pageUrl;
            options.allowMaximize = false;
            options.width = 450;
            options.height = 300;

            SP.UI.ModalDialog.showModalDialog(options);
        }

        filterCalendarView(ev: Event) {
            if (this.selectedId() == 0) {
                return true;
            }

            return ev.userId == this.selectedId();
        }

        getManagers(){
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

        getManagersSuccessCallback(){
            var listItemEnumerator = this.managersListItems.getEnumerator();

            while (listItemEnumerator.moveNext()) {
                var oListItem = listItemEnumerator.get_current();
                this.managment.push(new Manager(oListItem));
                this.managment = this.managment.sort((m1:Manager, m2:Manager) => { return m1.index - m2.index; });
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
                    basicWeek: { buttonText: "Неделя" }
                },
                locale: "ru",
                displayEventTime: false,
                events: this.events,
                eventClick: (event:any) => {
                    this.openDispForm(event);
                    return false;
                },
                eventRender: (event:any, element:any, view:any) => {
                    return this.filterCalendarView(event);
                },
            });

            $("#loader").removeClass("active");
            $(".kor-semantic").show();
        }

        errorCallback(sender:any, args:any) {
            alert('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        }
    }

    export class Manager implements IManager{
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
            $.model.selectedId(this.userId);
            $("#calendar").fullCalendar("rerenderEvents");
        }
    }

    export class Event {
        id: number;
        title: string;
        userId: number;//User
        start: string;//StartDate
        end: string;//EndDate
        status: string;
        color: string;

        constructor(oListItem: any) {
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
    }

    export class MockManager implements IManager{
        userId: number;
        name: string;
        imageUrl: string;
        position: string;
        index: number;
        isSelected = ko.observable();

        constructor(id: number, name: string, login: string, imageUrl: string, position: string) {
            this.filter = this.filter.bind(this);
            this.userId = id;
            this.name = name;
            this.imageUrl = imageUrl;
            this.isSelected(false);
            this.position = position;
            this.index = 0;
        }

        filter() {
            $.model.selectedId(this.userId);
            $("#calendar").fullCalendar("rerenderEvents");
        }
    }

    export interface IManager{
        userId: number;
        name: string;
        imageUrl: string;
        position: string;
        index: number;
        isSelected: any;
    }

}