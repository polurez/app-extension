import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'HelloWorldApplicationCustomizerStrings';

const LOG_SOURCE: string = 'HelloWorldApplicationCustomizer';
const VERSION = '1.0.0.29';
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

    //GET APP ID
    var appWebFullUrl = '';
    var webServerRelativeUrl = '';
    var a = window.location.pathname;
    var check = a.substring(1, a.indexOf('/Forms/AllItems.aspx'));

    if (check.indexOf('/') != -1) {
      webServerRelativeUrl = check.substring(0, check.lastIndexOf('/'));
      webServerRelativeUrl = '/' + webServerRelativeUrl;
    }
    const userCustomActionUrl = window.location.origin + webServerRelativeUrl + '/_api/Web/UserCustomActions';
    var xhrs = new XMLHttpRequest();
    xhrs.open('GET', userCustomActionUrl);
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
    }
    xhrs.send();
    //ADD CLICK HANDLER
    var addClickHandler = function () {
      setTimeout(function () {
        var aa = document.getElementsByTagName('a');
        for (let i = 0; i < aa.length; i++) {
          if (aa[i].href.indexOf('.mmap') == aa[i].href.length - 5) {
            aa[i].addEventListener("click", function (event) {
              event.preventDefault();
              event.stopPropagation();
              var path = this.href.toLowerCase();
              path = path.replace(window.location.origin, '');
              var webAbsoluteUrl = window.location.origin + webServerRelativeUrl

              if (webServerRelativeUrl === '/')
                webServerRelativeUrl = '';
              if ((webServerRelativeUrl.length > 1) && (webServerRelativeUrl.indexOf('/') == -1))
                webServerRelativeUrl = '../../../' + webServerRelativeUrl;

              var url = webServerRelativeUrl + '/_api/web/getfilebyserverrelativeurl(\'' + path + '\')/ListItemAllFields?$expand=ParentList';
              var xhr = new XMLHttpRequest();
              xhr.open('GET', url, false);
              xhr.setRequestHeader('Accept', 'application/json;odata=verbose');
              xhr.onload = function () {
                var temp = this as XMLHttpRequest;
                var data = JSON.parse(temp.response);
                var listId = data.d.ParentList.Id;
                var itemId = data.d.Id;
                if (appWebFullUrl.indexOf(webServerRelativeUrl + '/MindManagerSharePointApp') == -1) {
                  appWebFullUrl += webServerRelativeUrl + '/MindManagerSharePointApp';
                }
                window.location.href = appWebFullUrl + '/Pages/Default.aspx?SPHostUrl=' + encodeURI(webAbsoluteUrl) + '&SPAppWebUrl=' + encodeURI(appWebFullUrl) + '&SPListItemId=' + itemId + '&SPListId=%7B' + listId + '%7D';
              };
              xhr.send();
              return false;
            })
          }
        }
      }, 1000);
    }
    setInterval(function () {
      let isCalled = false;
      var imgs = document.querySelectorAll('img');
      for (var i = 0; i < imgs.length; i++) {
        if (((imgs[i].title.indexOf('mmap') + 1) || ((imgs[i].classList[0] == 'FileTypeIcon-icon') && (imgs[i].parentElement.title.indexOf('mmap') + 1))) && (imgs[i].src.indexOf('genericfile.png') == imgs[i].src.length - 15)) {
          imgs[i].src = imgBase64;
          if (!isCalled) isCalled = true;
        }
      }
      addClickHandler();
    }, 500);
    return Promise.resolve();
  }
}