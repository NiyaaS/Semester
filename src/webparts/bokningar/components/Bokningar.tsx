import * as React from 'react';
import styles from './Bokningar.module.scss';
import { IBokningarProps } from './IBokningarProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp, Items, Item } from "@pnp/sp";
import { CurrentUser } from '@pnp/sp/src/siteusers';
import { DefaultButton, PrimaryButton, TeachingBubbleBase, find } from 'office-ui-fabric-react';
import { Panel } from 'office-ui-fabric-react/lib/Panel';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { addDays, getDateRangeArray, isInDateRangeArray } from 'office-ui-fabric-react/lib/utilities/dateMath/DateMath';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { Calendar, DayOfWeek, DateRangeType } from 'office-ui-fabric-react/lib/Calendar';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';


const DayPickerStrings = {
  months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
  goToToday: 'Go to today',
  weekNumberFormatString: 'Week number {0}',
  prevMonthAriaLabel: 'Previous month',
  nextMonthAriaLabel: 'Next month',
  prevYearAriaLabel: 'Previous year',
  nextYearAriaLabel: 'Next year',
  prevYearRangeAriaLabel: 'Previous year range',
  nextYearRangeAriaLabel: 'Next year range',
  closeButtonAriaLabel: 'Close'
};

export interface IBokningarState {
  ItemsHolder: any;
  items: any;
  AnsvarigId: string;
  itemId: number;
  IsAdmin: boolean;
  showPanel: boolean;
  Title: string;
  Status: string;
  SlutDatum: Date | null;
  StartDatum: Date | null;
  ResDatum: Date[];
  Ansvarig: string;
  showStartCalendar: boolean;
  showSlutCalendar: boolean;
  buttonString: string;
  _buttonString: string;
  CurrentUserId: number;
  hideDialog: boolean;
}

export default class Bokningar extends React.Component<IBokningarProps, IBokningarState> {

  private _calendarButtonElement: HTMLElement;
  private calendarButtonElement: HTMLElement;


  private _columns: IColumn[];
  private options: IDropdownOption[];

  constructor(props: IBokningarProps, state: IBokningarState) {
    super(props);

    this.state = {
      ItemsHolder: [],
      items: [],
      AnsvarigId: '',
      itemId: null,
      IsAdmin: null,
      showPanel: false,
      Title: '',
      Status: '',
      StartDatum: null,
      SlutDatum: null,
      ResDatum: [],
      Ansvarig: '',
      showStartCalendar: false,
      showSlutCalendar: false,
      buttonString: 'Välj Start Datum',
      _buttonString: 'Välj Slut Datum',
      CurrentUserId: null,
      hideDialog: true,
    };

    this._columns = [
      { key: 'column1', name: 'Namn o Efternamn', fieldName: 'Title', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column2', name: 'Start Datum', fieldName: 'StartDatum', minWidth: 60, maxWidth: 80, isResizable: false, onRender: (item) => <span>{item.StartDatum.slice(0, 10)}</span> },
      { key: 'column2', name: 'Slut Datum', fieldName: 'SlutDatum', minWidth: 60, maxWidth: 80, isResizable: false, onRender: (item) => <span>{item.SlutDatum.slice(0, 10)}</span> },
      { key: 'column2', name: 'Status', fieldName: 'Status', minWidth: 50, maxWidth: 60, isResizable: false },
      { key: 'column2', name: 'Ansvarig', fieldName: 'Ansvarig', minWidth: 100, maxWidth: 200, isResizable: true, onRender: (item) => <span>{item.Ansvarig.Title}</span> },
      { key: 'column6', name: '', minWidth: 20, maxWidth: 50, isResizable: false, onRender: (item) => item.Status === "Skapad" || item.Status ==='Behandlas' ? <IconButton iconProps={{ iconName: 'Delete', }} title="Delete" ariaLabel="Delete" onClick={() => { this.deleteBookings(item.Id) }} /> : null },
      { key: 'column7', name: '', minWidth: 20, maxWidth: 50, isResizable: false, onRender: (item) => this.state.IsAdmin === false && ( item.Status === "Skapad" || item.Status ==='Behandlas') ? <IconButton iconProps={{ iconName: 'Sync' }} title="Uppdatera" ariaLabel="Uppdatera" onClick={() => { this.showPanelUppdate(item) }} /> : null },
      { key: 'column8', name: '', minWidth: 20, maxWidth: 50, isResizable: false, onRender: (item) => this.state.IsAdmin === true && (item.Status === "Skapad" || item.Status ==='Behandlas') ? <IconButton iconProps={{ iconName: 'Edit', }} title="Ändra Status" ariaLabel="Ändra Status" onClick={() => { this._showDialog(item.Id) }} /> : null },
    ];

    this.options = [
      //{ key: 'Skapad', text: 'Skapad' },
      { key: 'Beviljad', text: 'Beviljad' },
      { key: 'Behandlas', text: 'Behandlas' },
      { key: 'Avslagen', text: 'Avslagen' },
    ];
  }

  private _showDialog = (Id?:number): void => {
    this.setState({
      hideDialog: false,
      itemId: Id
    });
  }

  private _closeDialog = (): void => {
    this.setState({
      hideDialog: true,
    });
  }
  
  public componentDidMount() {
    this.getBookings();
  }

  private seSemester = (date: Date) => {

    let seSemesterbak = this.state.ItemsHolder.filter(x => Date.parse(x.StartDatum) >= date.valueOf());

    console.log(seSemesterbak);

    this.setState({
      items: seSemesterbak
    });
  }

  private addBooking = async (e: any) => {
    e.preventDefault();
  
    await sp.web.lists.getByTitle("Bookings").items.add({
      Title: this.state.Title,
      StartDatum: new Date(this.state.StartDatum.setHours(8)),
      SlutDatum: new Date(this.state.SlutDatum.setHours(17)),
      AnsvarigId: this.state.AnsvarigId
    })
    this.getBookings();
    this.hidePanel();
  }

  public getDatesInbetween = (StartDatum: string, SlutDatum: string) => {
    let currentDate = new Date(StartDatum);
    var resDates: Date[] = new Array();
    while(currentDate <= new Date(SlutDatum)){
        resDates.push(currentDate);
        currentDate = addDays(currentDate, +1);
    }
    return resDates;
  }

  private getBookings = async (today: Date = new Date()) => {
    // console.log(today);
    await sp.web.currentUser.get()
    .then((user: CurrentUser) => {
      //console.log(user);
      this.setState({
        CurrentUserId: user['Id'],
        Title: user['Title'],
        IsAdmin: user['IsSiteAdmin']
      });
    sp.web.lists.getByTitle('Bookings').items.select('*', 'Ansvarig/Title', 'Ansvarig/Id').expand('Ansvarig').get()
      .then((result: any) => {
        //console.log(result);
        let semesterBak: any = result.filter(x => Date.parse(x.StartDatum) >= Date.now().valueOf());
        //let allaBokadeDatum: Date[] = result.map([].concat(...result.map(f => this.getDatesInbetween(f.StartDatum, f.SlutDatum))));
        // let resDateUser: any = result.filter(x => x.AuthorId === this.state.CurrentUserId)
        
        this.setState({
          items: semesterBak,
          ItemsHolder: result,
          ResDatum:[].concat(...result.filter(x => x.AuthorId === this.state.CurrentUserId && new Date(x.StartDatum) > new Date()).map(f => this.getDatesInbetween(f.StartDatum, f.SlutDatum)))
        });
      })//.then(() => {console.log(this.state.ResDatum)})
    }); 
      // console.log(this.state.ResDatum);
  }

  private uppdateBooking = (id: number): void => {
    //console.log(id);
    sp.web.lists.getByTitle('Bookings').items.getById(id).update({
      StartDatum: new Date(this.state.StartDatum.setHours(8)),
      SlutDatum: new Date(this.state.SlutDatum.setHours(17)),
    }).then(() => {
      this.getBookings();
      this.hidePanel();
    });
  }

  private uppdateStatus = (id: number): void => {
    //console.log(id);
    sp.web.lists.getByTitle('Bookings').items.getById(id).update({
      Status: this.state.Status,
    }).then(() => {
      this.getBookings();
      this._closeDialog();
    });
  }

  private deleteBookings = (id): void => {

    sp.web.lists.getByTitle("Bookings").items.getById(id).delete()
      .then(() => {
        this.getBookings();
      });
  }

  private showPanel = (): void => {
    this.setState({
      showPanel: true,
    });

  }

  private showPanelUppdate = (item: any): void => {
    this.getBookings(new Date(item.StartDatum));
    //console.log(item.StartDatum.toLocaleDateString());
    console.log(item);
    this.setState({
      showPanel: true,
      StartDatum: new Date(item.StartDatum),
      SlutDatum: new Date(item.SlutDatum),
      itemId: item.Id,
    });
  }

  private hidePanel = (): void => {
    this.setState({
      showPanel: false,
      StartDatum: null,
      SlutDatum: null,
      itemId: null,
      AnsvarigId: '',
    });
  }

  private getPeoplePickerItems = (items: any[]) => {
    // console.log(items[0].id);
    this.setState({
      AnsvarigId: items[0].id
    });
  }

  private onSelectStartDate = (date: Date): void => {
    let isInRange = isInDateRangeArray(addDays(date, +1), this.state.ResDatum);
    //let dateRange = getDateRangeArray(this.state.StartDatum, DateRangeType, Date[])
    console.log(isInRange);
    this.setState({
      StartDatum: date,
      showStartCalendar: false,
    });
  }

  private onSelectSlutDate = (date: Date): void => {
    //let testDatum = this.getDatesInbetween(this.state.StartDatum.toISOString(), date.toISOString());
    //let findDatum =  testDatum.map(x => {return isInDateRangeArray(x, this.state.ResDatum)});
    //let ifDatumContain = findDatum.map(x => )
    
    this.setState({
      SlutDatum: date,
      showSlutCalendar: false,
    });
  }

  private onClickDate = (e: any): void => {
    this.setState({
      showStartCalendar: !this.state.showStartCalendar
    });
  }
  private _onClickDate = (e: any): void => {
    this.setState({
      showSlutCalendar: !this.state.showSlutCalendar
    });
  }

  private _dropdownChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
    this.setState({
      Status: option.text
    });
  }

  public render(): React.ReactElement<IBokningarProps> {

    let filterPerson: any[] = this.state.IsAdmin === true ? this.state.items.filter(x => x.AnsvarigId === this.state.CurrentUserId):
    this.state.items.filter(x => x.AuthorId === this.state.CurrentUserId);

    let showButtons: JSX.Element = this.state.itemId !== null ? 
    <PrimaryButton text="Uppdatera" type="submit" onClick={() => { this.uppdateBooking(this.state.itemId) }} />
    :<div>
      <PeoplePicker
      context={this.props.context}
      titleText="Ansvarig"
      personSelectionLimit={1}
      groupName={"semester Owners"}
      showtooltip={true}
      isRequired={true}
      disabled={false}
      selectedItems={this.getPeoplePickerItems}
      showHiddenInUI={false}
      principalTypes={[PrincipalType.User]}
      ensureUser={true}
      resolveDelay={1000}
    />
    <PrimaryButton text="Ansök" type="submit" onClick={this.addBooking} disabled={this.state.StartDatum === null || this.state.SlutDatum === null || this.state.AnsvarigId === '' ? true:false }/>
    </div> 
    
 

    return (
      <div className={styles.bokningar}>
        <h2>Semester Ansökningar</h2>

        <div>
          <DatePicker
            strings={DayPickerStrings}
            showWeekNumbers={true}
            firstWeekOfYear={1}
            onSelectDate={this.seSemester}
            showMonthPickerAsOverlay={true}
            placeholder="Se semester från datum..."
            ariaLabel="Select a date"
          />
        </div><br />

        <DetailsList
          items={filterPerson}
          columns={this._columns}
          checkboxVisibility={2}
        />

        <Dialog
          hidden={this.state.hideDialog}
          onDismiss={() => { this.setState({ hideDialog: true }) }}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: 'Välj Status',
          }}
          modalProps={{ isBlocking: false, styles: { main: { maxWidth: 450 } } }}
        >
          <Dropdown
            placeholder="Välj Status"
            label="Status"
            options={this.options}
            onChange={this._dropdownChange}
            defaultSelectedKey={this.state.Status}
          />
          <DialogFooter>
            <PrimaryButton onClick={() => { this.uppdateStatus(this.state.itemId) }} text="Save" />
            <DefaultButton onClick={this._closeDialog} text="Cancel" />
          </DialogFooter>
        </Dialog>

        <form>
          <Panel
            isOpen={this.state.showPanel}
            closeButtonAriaLabel="Close"
            isLightDismiss={true}
            headerText="Semester Ansökan"
            onDismiss={this.hidePanel}
          >
            <div>
              <div ref={calendarBtn => (this._calendarButtonElement = calendarBtn!)}>
                <DefaultButton
                  onClick={this.onClickDate}
                  text={!this.state.StartDatum ? this.state.buttonString : this.state.StartDatum.toLocaleDateString()}
                />
              </div>
              {this.state.showStartCalendar && (
                <Calendar
                  onSelectDate={this.onSelectStartDate}
                  value={this.state.StartDatum!}
                  strings={DayPickerStrings}
                  isMonthPickerVisible={false}
                  restrictedDates={this.state.ResDatum}
                  minDate={new Date()}
                />

              )}<br />

              <div ref={calendarBtn => (this.calendarButtonElement = calendarBtn!)}>
                <DefaultButton
                  onClick={this._onClickDate}
                  text={!this.state.SlutDatum ? this.state._buttonString : this.state.SlutDatum.toLocaleDateString()}
                />
              </div>
              {this.state.showSlutCalendar && (
                <Calendar
                  onSelectDate={this.onSelectSlutDate}
                  value={this.state.SlutDatum!}
                  strings={DayPickerStrings}
                  isMonthPickerVisible={false}
                  restrictedDates={this.state.ResDatum}
                  minDate={this.state.StartDatum}
                />
              )}
              {showButtons}
            </div>
          </Panel>
        </form>
        <PrimaryButton text="Ansök" type="submit" onClick={this.showPanel} />
      </div>
    );
  }
}
