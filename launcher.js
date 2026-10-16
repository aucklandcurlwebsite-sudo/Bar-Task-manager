'use strict';
const url=window.DAYBOOK_CONFIG?.appsScriptUrl;
if(typeof url==='string'&&/^https:\/\/script\.google\.com\/(?:macros\/s|a\/macros\/[A-Za-z0-9.-]+\/s)\/[A-Za-z0-9_-]+\/exec$/.test(url)&&!url.includes('YOUR_DEPLOYMENT_ID')){
  document.querySelector('#setup').remove();
  const link=document.querySelector('#standalone');link.href=url;link.hidden=false;
  const frame=document.createElement('iframe');frame.title='Daybook task manager — records saved to Google Drive';frame.src=url;frame.referrerPolicy='strict-origin';document.body.append(frame);
}
