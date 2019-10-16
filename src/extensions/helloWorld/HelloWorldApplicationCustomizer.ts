import { override } from '@microsoft/decorators';
import { Log, EventArgs } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import * as _ from 'lodash';
import * as strings from 'HelloWorldApplicationCustomizerStrings';

const LOG_SOURCE: string = 'HelloWorldApplicationCustomizer';
const VERSION = '1.0.0.50';
const imgBase64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAEnQAABJ0Ad5mH3gAAAAZdEVYdFNvZnR3YXJlAEFkb2JlIEltYWdlUmVhZHlxyWU8AAADKmlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNS42LWMxMzIgNzkuMTU5Mjg0LCAyMDE2LzA0LzE5LTEzOjEzOjQwICAgICAgICAiPiA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIiB4bWxuczpzdFJlZj0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlUmVmIyIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgQ0MgMjAxNS41IChNYWNpbnRvc2gpIiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOjhDQzkwNkE5ODFFMjExRTY5ODY0ODE0ODNCNTdGMjRFIiB4bXBNTTpEb2N1bWVudElEPSJ4bXAuZGlkOjhDQzkwNkFBODFFMjExRTY5ODY0ODE0ODNCNTdGMjRFIj4gPHhtcE1NOkRlcml2ZWRGcm9tIHN0UmVmOmluc3RhbmNlSUQ9InhtcC5paWQ6OENDOTA2QTc4MUUyMTFFNjk4NjQ4MTQ4M0I1N0YyNEUiIHN0UmVmOmRvY3VtZW50SUQ9InhtcC5kaWQ6OENDOTA2QTg4MUUyMTFFNjk4NjQ4MTQ4M0I1N0YyNEUiLz4gPC9yZGY6RGVzY3JpcHRpb24+IDwvcmRmOlJERj4gPC94OnhtcG1ldGE+IDw/eHBhY2tldCBlbmQ9InIiPz7WHR/pAAAEZklEQVRIS4VWS0xcZRT+7mPuDEyhLcVHGuPKkLhwg7GVEItpYjeGxEdqtQLGEJYa4wJdkBgTVujChQsXhBhNtNEoEXWBcaMLy7QkGBcmJCY+2qFCYRiew517//v7nXMHCjMFP3Lm3nv+/57Hd849P07PC5/axDpwXQ9wKHB5deE4jt7rRbF3Q1gkicXOToyZr5+v6e4O11qXmz0YFVcloS69dxAbB8bwXq+7kq4Z7vs/MIMr1lhG7jJajVgyYLTMKo2eP45Nr3uQDIAoTjA98RRK5WpNfxC5nEcHL35pJVqhyNKI/onlXeOK/cZTWGsRRgkunG3F9o6pae8gl/Xxw89/wnnypa/UgWbAlDXq9Ec3HgZxsL4ZYW7qmZqmEa+8+b1w4mrkGxsxNrYirFM2tuLalhRS0M1tsydblTRi+jgSMSn0hWtL+WDkcRqP4TJ6Q4Nvv38d97Y1qfHYWLz16mlsbMdKn+g+vLJYq9HR8CV6CaTn7OlUU8M3P/6N3/8oMwrg8tPtOPNIHpUwESZhWeD3Pr6FpmxtMzE9PY3FRXGa1rCvr08DkJZRPuvx7huduF0KEcYWFy+0YW3ToBqxsFWr13rMz89jYGAA/f398Dx2ZbKtjaONLO1Zjwfuy+Phh07g2fMnaTS5axCZTAab5SLGxsbQ3NysuoT960qaIkSj5X1457XH0NvTwjZMNO16+L6PX65exfDwMEqlEumM+REarK2tcVX226MdHMssKOGHd0tqRNDb24vJyUlMTU2ho6ODmiydMRtdPQQzhTnc+ndRI22gSIsJFAoz+riysoJyuayytLSkOvFPB5Z/bIs6/HPjJo0mWF4pqfEGiqgTn93dT+jj7OwshoaGMDg4iCiKWIwKPNaBhXZY9cZECoXryGazCFjIYnFBO6M+iziO8GhnJ0ZHR5HP51UnRZaM+UHxie1jIoOYsh/LLFh5fV0jEOfFhQX4nt+QhUR6/NSDGBkZqRVWmsdFpVKplYZFFnqCzMEMfpu7hvvbm7jZMhrgnraAHBf5YXnIBg7319FFdHV1YXx8HBMTE2htbWXwOU4EGRXKf4KfCgs6gzzP4qPPV3FzpZXjWGYr+yGw6OxYxbkzLdgJI54DnJaBBHWndt3d3SoHwH3OuUtfWIcn2ep6RAr4EqUln6NRj85ogrzH/Jrl4AmrhhnRJV/MN7u4vVzBr989V7PWiJdf/5YUGVackbSfDNB2IiAdUlgaTGKOBKNzR9olxyyOH/ORIz2+l7DHSa4uHo6Mz2F9/vJntlKhgaZAB5vMD0eL6yEMDVwW2WVHSAN57Di5D0Oh0mWRDU7lq6RSK3oAPo3/dWONB86lT+xOaElJAGkmnSP6/ZF/GhP6BPzPAB5finmK+WwKadntShXXJi/q+mFw46gKz2GqjkFCuhITs5djGpCzgTSoLkKG3WO45vmcS+BkrYZ8FnqPhhvwRPD5UuAZ5XlXpHN86lryLpqbWNikmt7nHGRYA93D944G8B8LKRSMhKRj2QAAAABJRU5ErkJggg==";

export interface IHelloWorldApplicationCustomizerProperties {
  testMessage: string;
}

export default class HelloWorldApplicationCustomizer
  extends BaseApplicationCustomizer<IHelloWorldApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    console.log('Plugin v', VERSION, ' is active');
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    let files;
    let isCalledIcon = false;
    let appWebFullUrl = '';
    let webServerRelativeUrl = '';
    
    function getSiteAndLibraryPath(){
      const anchor = "_spPageContextInfo"; 
      var scripts = document.getElementsByTagName("script");
      var result = '';

      for(var  i=0; i<scripts.length; i++){
        if(scripts[i].innerText.indexOf(anchor) != -1){
          result = scripts[i].innerText; 
          break;
        } 
      }

      return result = JSON.parse(result.substring(result.indexOf("{"), result.indexOf(";"))).listUrl;
    }

    function getSitePath(){
      let result = getSiteAndLibraryPath();
      
      return result.substring(0, result.lastIndexOf("/"));
    }

    function getFolderPath(){
      let h = document.location.href;
      let singleParam = h.indexOf("&") == -1;

      if ((h.indexOf("RootFolder=") == -1) && (h.indexOf("id=") == -1)) return "";


      let index = h.indexOf("id=") == -1 ? h.indexOf("RootFolder=") + 11 : h.indexOf("id=") + 3;
      let path = singleParam ? decodeURIComponent(h.substring(index)) : decodeURIComponent(h.substring(index, h.indexOf("&")));

      return path.replace(getSiteAndLibraryPath(), "") + "/";
    }

    function addButtonEventListener(){
      var buttons = document.getElementsByTagName("button");
      _.each(buttons, b => {
        if (b.innerText &&  b.innerText.indexOf('.mmap') == b.innerText.length -5) b.addEventListener("click", buttonHandler);
      });
    }

    function initAppUrl() {
      const userCustomActionUrl = window.location.origin + getSitePath() + '/_api/Web/UserCustomActions';
      var xhrs = new XMLHttpRequest();
      xhrs.open('GET', userCustomActionUrl, false);
      xhrs.setRequestHeader('Accept', 'application/json;odata=verbose');
      xhrs.onload = function () {
        var temp = this as XMLHttpRequest;
        var data = JSON.parse(temp.response);
        for (var i = 0; i < data.d.results.length; i++) {
          if ((data.d.results[i].Description) && (data.d.results[i].Description.indexOf('mmapExtensionHandler') + 1)) {
            appWebFullUrl = data.d.results[i].Description.substring(data.d.results[i].Description.indexOf('?') + 1, data.d.results[i].Description.length);
            break;
          }
        }
      };
      xhrs.send();
    } 

    function buttonHandler(event){    
      event.preventDefault();
      event.stopPropagation();

      var fileName = event.target.innerText;

      let mmapProps = _.find(files, e => e.FileLeafRef == fileName);

      var pathToFile =  mmapProps == undefined 
      ? (getSiteAndLibraryPath()+ '/'+ getFolderPath() + fileName).toLowerCase()
      : mmapProps.FileRef.toLowerCase();

      webServerRelativeUrl = getSitePath();
      var webAbsoluteUrl = window.location.origin + getSitePath();
      
      /*valid url
      
      https://smdevelop-54018cb903bc12.sharepoint.com/sites/fa/MindManagerSharePointApp/Pages/Default.aspx
      ?SPHostUrl=https://smdevelop.sharepoint.com/sites/fa
      &SPAppWebUrl=https://smdevelop-54018cb903bc12.sharepoint.com/sites/fa/MindManagerSharePointApp
      &SPListItemId=88&SPListId=%7Bf35e8304-f748-4dce-8193-7008ca290923%7D
      */

      var url = webServerRelativeUrl + '/_api/web/getfilebyserverrelativeurl(\'' + pathToFile + '\')/ListItemAllFields?$expand=ParentList';
      var xhr = new XMLHttpRequest();
      xhr.open('GET', url);
      xhr.setRequestHeader('Accept', 'application/json;odata=verbose');
      xhr.onload = function () {
        var temp = this as XMLHttpRequest;
        var data = JSON.parse(temp.response);
        var listId = data.d.ParentList.Id;
        var itemId = data.d.Id;
        if (appWebFullUrl.indexOf(webServerRelativeUrl + '/MindManagerSharePointApp') == -1) {
          appWebFullUrl += webServerRelativeUrl + '/MindManagerSharePointApp';
        }
        let newTabUrl = appWebFullUrl + '/Pages/Default.aspx?SPHostUrl=' + encodeURI(webAbsoluteUrl) + '&SPAppWebUrl=' + encodeURI(appWebFullUrl) + '&SPListItemId=' + itemId + '&SPListId=%7B' + listId + '%7D';
        
        window.open(newTabUrl, '_blank');
      };
      xhr.send();

      return false;
    }

    function init() {
      isCalledIcon = false;
      initAppUrl();
      document.addEventListener("click", searchResultClickHandler);

      var a = window.location.pathname;
      var check = a.substring(1, a.indexOf('/Forms/'));

      if (check.indexOf('/') != -1) {
        webServerRelativeUrl = check.substring(0, check.lastIndexOf('/'));
        webServerRelativeUrl = '/' + webServerRelativeUrl;
      }
      let temp;
      var scripts = document.getElementsByTagName("script");
      _.forEach(scripts, element => {
        if(element.innerHTML.toString().indexOf("g_listData") != -1) temp = element.innerHTML;
      });
      
      let substr = temp.substring(temp.indexOf("Row") + 6, temp.indexOf("FirstRow") -2);
      files = JSON.parse(substr);
      console.log("parsed(substr) :", JSON.parse(substr));

      setInterval(wrapper, 100);
    }

    function searchResultClickHandler(event){
      if(event.target.href && (event.target.href.indexOf('DispForm') !== -1)){
        event.preventDefault();
        event.stopPropagation();

        let xhr = new XMLHttpRequest();
        xhr.open('GET', event.target.href);
        xhr.setRequestHeader('Accept', 'application/json;odata=verbose');
        xhr.onload = function () {
          let temp = this as XMLHttpRequest;
          
          let t = temp.response.substring(temp.response.indexOf("\"item\""));
          let item = JSON.parse(t.substring(8, t.indexOf('}') + 1));
          let ctx = temp.response.substring(temp.response.indexOf('_spPageContextInfo'));
          let _spPageContexInfo = JSON.parse(ctx.substring(19, ctx.indexOf('};') + 1));

          //init params for final link
          let SPHostUrl = _spPageContexInfo.webAbsoluteUrl;
          let SPAppWebUrl = _getAppFullUrl(SPHostUrl) + _spPageContexInfo.webServerRelativeUrl + "/MindManagerSharePointApp";
          let SPListItemId = item.Id;
          let SPListId = _spPageContexInfo.listId.substring(1, _spPageContexInfo.listId.length - 1);

          window.location.href = SPAppWebUrl + '/Pages/Default.aspx?SPHostUrl=' + encodeURI(SPHostUrl) + '&SPAppWebUrl=' + encodeURI(SPAppWebUrl) + '&SPListItemId=' + SPListItemId + '&SPListId=%7B' + SPListId + '%7D';
        };
        xhr.send();
      }
    }

    /**
     * @function addIcon
     * @description Change default icon to MindManager icon in all mmap files
     */
    function addIcon() {
      if (isCalledIcon) return;

      var imgs = document.querySelectorAll('img');

      for (var i = 0; i < imgs.length; i++) {
        if (_isValidImg(imgs[i])) {
          imgs[i].src = imgBase64;
          isCalledIcon = true;
        }
      }

      isCalledIcon = false;
    }

    /**
     * @function wrapper 
     * @description Function for wrapping two another function for calling with same interval and right order 
     */
    function wrapper() {
      addButtonEventListener();
      addIcon();
    }

    /**
     * @function _isValidImg
     * @description Check is the image mmap icon
     * @param {Object} img  object for checking  
     * @return {Boolean} Return true if img is mmap icon
     */
    function _isValidImg(img) {
      // Already setted icon
      if(img.src === imgBase64) return true; 

      //Check First parent Node
      if(img.parentElement.title && (img.parentElement.title.indexOf('mmap') !== -1)) return true;
      if(img.parentElement.attributes.getNamedItem('aria-label') && (img.parentElement.attributes.getNamedItem('aria-label').value.indexOf('mmap') !== -1)) return true;

      //CheckSecond parent Node
      if(img.parentElement.parentElement.innerText && (img.parentElement.parentElement.innerText.indexOf('mmap') !== -1)) return true;

      //Check for old layout (third parent Node)
      let oldAriaLabel = img.parentElement.parentElement.parentElement.attributes.getNamedItem('aria-label');
      let oldTitle = img.parentElement.parentElement.parentElement.title;

      if(oldAriaLabel && oldAriaLabel.value.indexOf('mmap') !== -1) return true;
      if(oldTitle && oldTitle.indexOf('mmap') !== -1) return true;

      //Check for Search results 
      if(img.src.substring(img.src.indexOf('?') - 7, img.src.indexOf('?')) === 'spo.svg') {
        //init p variable for 3-rd parent element shortcut
        let p = img.parentElement.parentElement.parentElement;
        if(p && p.href && p.className && p.className.indexOf('FileSearchResult') !== -1 && p.href.indexOf('?ID=') !== -1 ) return true;
      }  

      //Check for swedish layout (fourth parent Node)
      let ariaLabel = img.parentElement.parentElement.parentElement.parentElement.attributes.getNamedItem('aria-label');
      let title = img.parentElement.parentElement.parentElement.parentElement.title;

      if(ariaLabel && ariaLabel.value.indexOf('mmap') !== -1) return true;
      if(title && title.indexOf('mmap') !== -1) return true;

      //Check Tile View (sixth parent)
      if(img.parentElement.parentElement.parentElement.className === 'ItemTile-fileIconContainer') {
        let parent = img.parentElement.parentElement.parentElement.parentElement.parentElement.parentElement;
        if(parent && (parent.href) && (parent.href.substring(parent.href.length - 4) == 'mmap')) return true;
      }

      return false;
    }

    function _getAppFullUrl(absoluteUrl) {
      let userCustomActionUrl = absoluteUrl + '/_api/Web/UserCustomActions';
      let xhr = new XMLHttpRequest();

      xhr.open('GET', userCustomActionUrl, false);
      xhr.setRequestHeader('Accept', 'application/json;odata=verbose');
      xhr.send();

      if (xhr.status === 200) {
        let data = JSON.parse(xhr.response);

        for (let i = 0; i < data.d.results.length; i++) {
          if ((data.d.results[i].Description) && (data.d.results[i].Description.indexOf('mmapExtensionHandler') + 1)) {
            return data.d.results[i].Description.substring(data.d.results[i].Description.indexOf('?') + 1, data.d.results[i].Description.length);
          }
        }
      }
    }

      init();

    return Promise.resolve();
  }
}