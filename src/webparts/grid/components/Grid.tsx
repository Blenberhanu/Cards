import * as React from 'react';
import styles from './Grid.module.scss';
import { IGridProps, IGridState, ISPDocument, ISPDocuments, IGridItem } from './Grid.types';
import { ISize } from 'office-ui-fabric-react/lib/Utilities';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import ReactPagination from "react-js-pagination";
import "bootstrap/dist/css/bootstrap.min.css";
// Used to render document cards
import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardDetails,
  DocumentCardTitle,
  IDocumentCardPreviewProps,
  DocumentCardLocation,
  DocumentCardType
} from 'office-ui-fabric-react/lib/DocumentCard';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import * as jquery from 'jquery';
import { GridLayout } from '../../../components/gridLayout/index';
import GridWebPart from '../GridWebPart';
//import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import {
  PrimaryButton, DefaultButton, IconButton
} from 'office-ui-fabric-react/lib/Button';
//import FileViewer from 'react-file-viewer';
//import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import * as CSS from 'csstype';
var divStyle: CSS.Properties = {
  minHeight: 'auto'
};
let myfile: string = '';
let mytype: string = '';
let currentList = [];
var _categoryitems = [];
var _isParentCategory = true;
var hasFilledCategories = false;
var isSubCategory = false;
var hasFilledSubCategory = false
var _selectedCategory = "All"
var _currentParrent = ""
//import { useId, useBoolean } from '@uifabric/react-hooks';
import { IDragOptions, Modal } from 'office-ui-fabric-react/lib/Modal';
import { ContextualMenu } from 'office-ui-fabric-react/lib/ContextualMenu';
import { IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
import { FontWeights, getTheme, mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { Dropdown, IDropdownOption, IDropdownStyles, TextField } from 'office-ui-fabric-react';
//const [isModalOpen, { setTrue: showModal, setFalse: hideModal }] = useBoolean(false);
//const [isDraggable, { toggle: toggleIsDraggable }] = useBoolean(false);
export default class Grid extends React.Component<IGridProps, IGridState> {

  constructor(props: IGridProps) {
    super(props);
    if (Environment.type === EnvironmentType.Local) {

      this.state = {
        CardsData: [{
          thumbnail: "https://lorempixel.com/400/200/technics/1/",
          title: "Adventures in SPFx",
          name: "Perry Losselyong",
          profileImageSrc: "https://robohash.org/blanditiisadlabore.png?size=50x50&set=set1",
          location: "SharePoint",
          activity: "3/13/2019"
        }, {
          thumbnail: "https://lorempixel.com/400/200/technics/2",
          title: "The Wild, Untold Story of SharePoint!",
          name: "Ebonee Gallyhaock",
          profileImageSrc: "https://robohash.org/delectusetcorporis.bmp?size=50x50&set=set1",
          location: "SharePoint",
          activity: "6/29/2019"
        }, {
          thumbnail: "https://lorempixel.com/400/200/technics/3",
          title: "Low Code Solutions: PowerApps",
          name: "Seward Keith",
          profileImageSrc: "https://robohash.org/asperioresautquasi.jpg?size=50x50&set=set1",
          location: "PowerApps",
          activity: "12/31/2018"
        }, {
          thumbnail: "https://lorempixel.com/400/200/technics/4",
          title: "Not Your Grandpa's SharePoint",
          name: "Sharona Selkirk",
          profileImageSrc: "https://robohash.org/velnammolestiae.png?size=50x50&set=set1",
          location: "SharePoint",
          activity: "11/20/2018"
        }, {
          thumbnail: "https://lorempixel.com/400/200/technics/5/",
          title: "Get with the Flow",
          name: "Boyce Batstone",
          profileImageSrc: "https://robohash.org/nulladistinctiomollitia.jpg?size=50x50&set=set1",
          location: "Flow",
          activity: "5/26/2019"
        }],
        currentPage: 1,
        CardsDataPerPage: 8,
        hideDialog: false,
        file: '',
        type: 'csv',
        CardImagelink: '',
        cardTitle: '',
        downloadPPT: '',
        downloadPdf: '',
        Id: '',
        Categoryitems: _categoryitems
      }
    }
    else if (Environment.type == EnvironmentType.SharePoint ||
      Environment.type == EnvironmentType.ClassicSharePoint) {

      this._getParentCategories()
      this.state = {
        CardsData: this.Libitems,
        currentPage: 1,
        CardsDataPerPage: 6,
        hideDialog: false,
        file: '',
        type: 'csv',
        CardImagelink: '',
        cardTitle: '',
        downloadPPT: '',
        downloadPdf: '',
        Id: '',
        Categoryitems: _categoryitems
      }
      currentList = this.state.CardsData;
    }

    this.handleClick = this.handleClick.bind(this);
    this.handleChange = this.handleChange.bind(this);

  }
  private Libitems = [] as any;
  private currentSelection;

  handleClick(event) {


    this.setState({
      currentPage: Number(event)
    });

  }


  private _CustomSearch(e): Promise<ISPDocuments> {
    var result = [];
    var searchTerm = e.target.value;
    _isParentCategory = false
    if (e.key === 'Enter') {
      if (e.target.value == '') {
        this.setState({ CardsData: currentList });
        this.render();
        return
      }
      jquery.ajax({
        url: `/_api/search/query?querytext='` + searchTerm + `'&amp;sourceid='e7ec8cee-ded8-43c9-beb5-436b54b31e84'&amp;selectproperties='ows_Description,ows_SiteDescription,ows_Description'`,
        type: "GET",
        async: false,
        headers: { 'Accept': 'application/json; odata=verbose;' },
        success: (resultData) => {
          var item = {};
          resultData.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results.forEach(row => {
            row.Cells.results.forEach(function (child) {
              item[child.Key] = {
                Value: child.Value,
                ValueType: child.ValueType
              };
            });
            // add result object to result array
            let url = "/sites/IPSteeringGroup" + item["Path"].Value.substring(item["SiteName"].Value.length, item["Path"].Value.length); //data[index].ServerRelativeUrl;
            var fileDate = new Date(item["LastModifiedTime"].Value).toLocaleDateString('en-US', {
              day: 'numeric',
              month: 'numeric',
              year: 'numeric',
            })
            var cardItem = {
              thumbnail: this.getPreviewUrl(url),
              title: item["Title"].Value,
              name: item["Title"].Value.split(".")[0] + "." + item["FileExtension"].Value,
              profileImageSrc: item["SiteDescription"].Value,
              location: item["OriginalPath"].Value,
              activity: fileDate,
              Id: ''
            }
            if (item["FileExtension"].Value == "ppt" || item["FileExtension"].Value == "pptx")
              result.push(cardItem)
          });

        },
        error: function (jqXHR, textStatus, errorThrown) {
          return null
        }
      });
      this.setState({ CardsData: result });
      this.render();

      this.Libitems = result
      return null
    }
  }
  public getPreviewUrl(path) {
    var _url = this.gethostName() + `/_api/v2.0/sharePoint:` + path + `:/driveItem/thumbnails/0/c1280x720`
    var resultUrl;
    jquery.ajax({
      url: _url,
      type: "GET",
      async: false,
      headers: {
        "Content-Type": "application/json",
        "Accept": "application/json"
      },
      success: (resultData) => {

        resultUrl = resultData.url;
      },
      error: function (jqXHR, textStatus, errorThrown) {
        return null
      }
    });
    return resultUrl;

  }
  public _getListData(e): Promise<ISPDocuments> {
    var reactHandler = this;
    var result = [];
    var test = e;

    var siteUrl = this.props.siteurl
    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/GetFolderByServerRelativeUrl('Shared%20Documents')/Files?$select=ListItemAllFields/Category,ListItemAllFields/IsParent,ListItemAllFields/SubCategory,ID,TimeLastModified,ListItemAllFields/File_x0020_Type,ServerRelativeUrl,FileRef,Name,DocumentType,EncodedAbsUrl,EncodedAbsThumbnailUrl,CreatedBy/Title&$expand=ListItemAllFields,File,CreatedBy&$orderby=ListItemAllFields/Category asc`,
      type: "GET",
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: (resultData) => {
        _categoryitems = []
        _categoryitems.push({ key: "Go To Parent List", text: "Go To Parent List" })
        var data = resultData.d.results
        for (let index = 0; index < data.length; index++) {
          let url = data[index].ServerRelativeUrl;
          var fileDate = new Date(data[index].TimeLastModified).toLocaleDateString('en-US', {
            day: 'numeric',
            month: 'numeric',
            year: 'numeric',
          })
          var filter = data[index].ListItemAllFields.Category
          var parent = data[index].ListItemAllFields.Category
          _selectedCategory = _currentParrent;
          if (data[index].ListItemAllFields.IsParent) {
            continue;
          }
          if (_categoryitems.filter(ev => ev.key === data[index].ListItemAllFields.SubCategory).length == 0) {
            _categoryitems.push({ key: data[index].ListItemAllFields.SubCategory, text: data[index].ListItemAllFields.SubCategory })

          }
          if (!_isParentCategory) {
            filter = data[index].ListItemAllFields.SubCategory
          }
          else {
            //filter = data[index].ListItemAllFields.SubCategory
          }
          if (filter != e && !isSubCategory)
            continue

          if (_currentParrent != parent && isSubCategory) {
            continue;
          }
          if (filter != e)
            continue

          var Ditem = {
            thumbnail: this.getPreviewUrl(data[index].ServerRelativeUrl),// 
            title: data[index].ListItemAllFields.Description,
            name: data[index].Name,
            profileImageSrc: data[index].ListItemAllFields.Category,
            location: url,
            activity: fileDate,
            Id: data[index].ID
          }
          result.push(Ditem)
        }
        if (_isParentCategory && !hasFilledSubCategory) {
          hasFilledSubCategory = true
        }
      },
      error: function (jqXHR, textStatus, errorThrown) {
        return null
      }
    });
    this.Libitems = result
    currentList = result;
    return test

  }

  public _getParentCategories(): Promise<ISPDocuments> {
    var reactHandler = this;
    var result = [];
    var test;
    var siteUrl = this.props.siteurl
    if (!hasFilledCategories) {
      _categoryitems = []
      _categoryitems.push({ key: "All", text: "All", isSelected: true })
    }
    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/GetFolderByServerRelativeUrl('Shared%20Documents')/Files?$select=ListItemAllFields/Category,
      ListItemAllFields/Title,ListItemAllFields/IsParent,ListItemAllFields/Title,ID,
      TimeLastModified,ListItemAllFields/File_x0020_Type,ServerRelativeUrl,FileRef,Name,DocumentType,EncodedAbsUrl,EncodedAbsThumbnailUrl,CreatedBy/Title&$expand=ListItemAllFields,File,CreatedBy&$orderby=Name asc`,
      type: "GET",
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: (resultData) => {
        var data = resultData.d.results
        for (let index = 0; index < data.length; index++) {
          let url = data[index].ServerRelativeUrl;
          var fileDate = new Date(data[index].TimeLastModified).toLocaleDateString('en-US', {
            day: 'numeric',
            month: 'numeric',
            year: 'numeric',
          })
          if (!data[index].ListItemAllFields.IsParent) {
            continue;
          }
          if (!hasFilledCategories)
            _categoryitems.push({ key: data[index].ListItemAllFields.Category, text: data[index].ListItemAllFields.Category })

          var Ditem = {
            thumbnail: this.getPreviewUrl(data[index].ServerRelativeUrl),// 
            title: data[index].ListItemAllFields.Title,
            name: data[index].Name,
            profileImageSrc: data[index].ListItemAllFields.Category,
            location: url,
            activity: fileDate,
            Id: data[index].ID
          }
          result.push(Ditem)
        }
        hasFilledCategories = true
      },
      error: function (jqXHR, textStatus, errorThrown) {
        return null
      }
    });
    this.Libitems = result
    currentList = result;
    _isParentCategory = true
    return test

  }

  gethostName() {
    var parser = document.createElement('a');
    parser.href = this.props.siteurl;
    return parser.origin;

  }
  private _showDialog = (e, data): void => {

    _selectedCategory = data.profileImageSrc
    if (_isParentCategory) {
      _currentParrent = _selectedCategory
      this._getListData(data.profileImageSrc)
      this.setState({
        CardsData: currentList, Categoryitems: _categoryitems
      })
      this.render()
      _isParentCategory = false
      isSubCategory = true;
    }
    // else if (isSubCategory) {
    //   _selectedCategory = "Sub Category 1"//data.profileImageSrc
    //   this._getListData(_selectedCategory)
    //   this.setState({
    //     CardsData: currentList, Categoryitems: _categoryitems
    //   })
    //   this.render()
    //   _isParentCategory = false
    // }
    else {
      this.currentSelection = data;
      this.setState({
        hideDialog: true, downloadPPT: data.location, CardImagelink: data.thumbnail,
        downloadPdf: data.location, file: data.name, Categoryitems: _categoryitems
      })
      this.render()
    }
  }

  private _closeDialog = (): void => {

    this.setState({ hideDialog: false });
    this.render();
  }

  // private _getPDF(Id) {

  //   var fileType = Id.split('.')[1]
  //   if (fileType == "ppt") {
  //     fileType += "x"
  //   }
  //   const body: string = JSON.stringify({
  //     'ItemId': "/Documents/" + Id,
  //     type: fileType,
  //     name: Id.split('.')[0]
  //   });

  //   var msFlowUrl = "https://prod-103.westeurope.logic.azure.com:443/workflows/d90d4a436c4f421bb2e4f1289670d45a/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=JpaFTvGLhH8WHMUWWJu4qwUbfCDGiijvUmZAgSJgFv8"
  //   jquery.ajax({
  //     url: msFlowUrl,
  //     type: "POST",
  //     data: body,
  //     async: false,
  //     processData: false,
  //     encoding: null,
  //     headers: {
  //       'Accept': 'application/json; odata=verbose;',
  //       'Content-type': 'application/json'
  //     },
  //     beforeSend: function (request) {
  //       request.overrideMimeType('text/plain; charset=x-user-defined');
  //     },
  //     success: function (response) {
  //       var binary = "";
  //       var responseTextLen = response.length;

  //       for (var i = 0; i < responseTextLen; i++) {
  //         binary += String.fromCharCode(response.charCodeAt(i) & 255)
  //       }

  //       var a = document.createElement('a');
  //       a.href = "data:application/pdf;base64," + btoa(binary);
  //       a.download = Id.split('.')[0] + '.pdf';
  //       document.body.appendChild(a);
  //       a.click();
  //       a.remove();
  //     },
  //     error: function (jqXHR, textStatus, errorThrown) {
  //       console.log(jqXHR)
  //     }
  //   });

  // }
  onDropdownChange = (event: React.FormEvent<HTMLDivElement> | any, option: any = {}, index?: number) => {

    if ((option.key == "All" || option.key == "Go To Parent List") && !_isParentCategory) {
      hasFilledCategories = false
      this._getParentCategories()
      _selectedCategory = "All"
      this.setState({
        CardsData: currentList,
      })
      this.render()
      isSubCategory = false

    } else {
      if (option.key != "All" && !isSubCategory) {
        _currentParrent = option.key
        _selectedCategory = _currentParrent
        this._getListData(option.key)
        this.setState({
          CardsData: currentList,
        })
        this.render()
        _isParentCategory = false;
        isSubCategory = true
      }
      else if (isSubCategory && option.key != "Go To Parent List") {
        _currentParrent = _selectedCategory
        this._getListData(option.key)
        this.setState({
          CardsData: currentList,
        })
        this.render()
        _isParentCategory = false;
        hasFilledSubCategory = true
      }
    }
    //}
  }
  public render(): React.ReactElement<IGridProps> {

    let pagedItems: any[] = this.state.CardsData;
    const totalItems: number = pagedItems.length;
    let showPages: boolean = false;
    const maxEvents: number = 6; // Use any page size you want
    const { currentPage } = this.state;

    if (true && totalItems > 0 && totalItems > maxEvents) {
      // calculate the page size
      const pageStartAt: number = maxEvents * (currentPage - 1);
      const pageEndAt: number = (maxEvents * currentPage);

      pagedItems = pagedItems.slice(pageStartAt, pageEndAt);
      showPages = true;
    }
    const titleId = 'TestModal';
    return (

      <>


        <div className={styles.grid}>

          <TextField label="Search" onKeyPress={((e) => this._CustomSearch(e))} />
          <Stack tokens={stackTokens}>
            <Dropdown placeholder="All" onChange={this.onDropdownChange} label={"Category: " + _selectedCategory} options={_categoryitems} styles={dropdownStyles} />
          </Stack>
          <GridLayout
            items={pagedItems}
            onRenderGridItem={(item: any, finalSize: ISize, isCompact: boolean) => this._onRenderGridItem(item, finalSize, isCompact)} />

          <div>
            <Modal isOpen={this.state.hideDialog} onDismiss={this._closeDialog} isBlocking={false}
              containerClassName={contentStyles.container}
            >
              <div className={contentStyles.header}>
                <span id='Tsets' >{this.state.cardTitle}</span>
                <IconButton
                  styles={iconButtonStyles}
                  iconProps={cancelIcon}
                  ariaLabel="Close popup modal"
                  onClick={this._closeDialog}
                />
              </div>
              <div className={contentStyles.body}>
                <div className='col-md-12'>

                  {/* <div className='col-md-12'> */}
                  <img className={contentStyles.img} width="100%" src={this.state.CardImagelink} />
                  {/* </div> */}

                  {/* <FileViewer
                fileType={this.state.type}
                filePath={this.state.file}
              /> */}
                  <div className='row'>
                    <div className="card-body text-center">
                      <h6 className="card-title">Download as</h6>
                      <div className="btn-group">
                        {/* <a href="#" onClick={((e) => this._getPDF(this.state.file))} className="btn btn-default stretched-link"> PDF</a> */}
                        <a href={this.state.downloadPPT} target="_blank" className="btn btn-primary stretched-link" download={this.state.cardTitle}> PPT</a>

                      </div>

                    </div>
                  </div>
                </div>
                {/* <div className={contentStyles.button}>
                  <PrimaryButton text="Close Window" onClick={this._closeDialog} />
                </div> */}


              </div>
            </Modal>

            {showPages &&
              <ReactPagination
                itemClass="page-item"
                linkClass="page-link"
                hideNavigation
                activePage={currentPage}
                itemsCountPerPage={maxEvents}
                totalItemsCount={totalItems}
                pageRangeDisplayed={8}
                onChange={this.handleClick.bind(this)}
              />}
          </div>
        </div>


      </>

    );
  }

  private _onRenderGridItem = (item: any, finalSize: ISize, isCompact: boolean): JSX.Element => {
    const previewProps: IDocumentCardPreviewProps = {
      previewImages: [
        {
          previewImageSrc: item.thumbnail,
          imageFit: ImageFit.cover,
          height: 130
        }
      ]
    };
    var FileType = item.name.split('.')[1]
    return <div
      className={styles.documentTile}
      data-is-focusable={true}
      role="listitem"
      aria-label={item.title}
    >
      <DocumentCard
        type={isCompact ? DocumentCardType.compact : DocumentCardType.normal}
        onClick={((e) => this._showDialog(e, item))}
      // onClick={showModal}
      >

        <DocumentCardPreview {...previewProps} />
        {/* <FileViewer {...previewProps}
          fileType={FileType}
          filePath={item.location}
        /> */}
        {/* {!isCompact && <DocumentCardLocation location={item.profileImageSrc} />} */}
        <DocumentCardDetails>
          <DocumentCardTitle
            title={item.title}
            shouldTruncate={false}
          />
          {/* { <DocumentCardActivity
          //activity={item.activity}
          //people={[{ name: item.name, profileImageSrc: item.profileImageSrc }]}
          // people={[{ name: item.name, profileImageSrc: <FileTypeIcon type={IconType.image} size={ImageSize.small} application={ApplicationType.Excel} /> }]}
          />} */}

        </DocumentCardDetails>
      </DocumentCard>
    </div>;
  }

  handleChange(e) {
    // Variable to hold the original version of the list

    // Variable to hold the filtered list before putting into state
    let newList = [];

    // If the search bar isn't empty
    if (e.target.value !== "") {
      // Assign the original list to currentList


      // Use .filter() to determine which items should be displayed
      // based on the search terms
      newList = currentList.filter(item => {
        // change current item to lowercase
        const lc = item.name.toLowerCase();
        // change search term to lowercase
        const filter = e.target.value.toLowerCase();
        // check to see if the current list item includes the search term
        // If it does, it will be added to newList. Using lowercase eliminates
        // issues with capitalization in search terms and search content
        return lc.includes(filter);
      });
    } else {
      // If the search bar is empty, set newList to original task list
      newList = currentList;
    }
    // Set the filtered state based on what our rules added to newList
    this.setState({ CardsData: newList });
    this.render();
  }
}
const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300, float: 'right', marginBottom: '5px', marginTop: '5px' }
};
const options: IDropdownOption[] = [
  { key: 'Select Category', text: 'Select Category' },

];
//const stackTokens: IStackTokens = { childrenGap: 20 };



const dragOptions: IDragOptions = {
  moveMenuItemText: 'Move',
  closeMenuItemText: 'Close',
  menu: ContextualMenu,
};

const cancelIcon: IIconProps = { iconName: 'Cancel' };
const stackTokens: IStackTokens = { childrenGap: 1 };
const theme = getTheme();
const contentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    flexFlow: 'column nowrap',
    alignItems: 'stretch',
    width: '125vh',
    overflowY: 'hidden'
  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      borderTop: `4px solid ${theme.palette.themePrimary}`,
      color: theme.palette.neutralPrimary,
      display: 'flex',
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      padding: '12px 12px 14px 24px',
    },
  ],
  body: {
    flex: '4 4 auto',
    padding: '0 24px 24px 24px',
    overflowY: 'hidden',
    selectors: {
      p: { margin: '14px 0' },
      'p:first-child': { marginTop: 0 },
      'p:last-child': { marginBottom: 0 },
    },
  },
  button: {
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    //backgroundColor: "blue",
    height: "74.5px"

  }, img: {
    // margin: '10px auto 20px',
    display: 'block',
    //maxHeight: '50.5vh',

  }
});
const getStyles = () => {
  return {
    root: {
      maxWidth: '600px',

    }
  }
};

const iconButtonStyles = {
  root: {
    color: theme.palette.neutralPrimary,
    marginLeft: 'auto',
    marginTop: '4px',
    marginRight: '2px',
  },
  rootHovered: {
    color: theme.palette.neutralDark,
  },
};
