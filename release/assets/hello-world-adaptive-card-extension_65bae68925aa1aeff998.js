define("717e8239-32cd-4afc-af51-216d2d1dfc41_0.0.1",["@microsoft/sp-adaptive-card-extension-base"],function(e){return function(e){function t(t){for(var n,o,i=t[0],a=t[1],u=0,s=[];u<i.length;u++)o=i[u],Object.prototype.hasOwnProperty.call(r,o)&&r[o]&&s.push(r[o][0]),r[o]=0;for(n in a)Object.prototype.hasOwnProperty.call(a,n)&&(e[n]=a[n]);for(c&&c(t);s.length;)s.shift()()}var n={},r={1:0};function o(t){if(n[t])return n[t].exports;var r=n[t]={i:t,l:!1,exports:{}};return e[t].call(r.exports,r,r.exports,o),r.l=!0,r.exports}o.e=function(e){var t=[],n=r[e];if(0!==n)if(n)t.push(n[2]);else{var i=new Promise(function(t,o){n=r[e]=[t,o]});t.push(n[2]=i);var a,u=document.createElement("script");u.charset="utf-8",u.timeout=120,o.nc&&u.setAttribute("nonce",o.nc),u.src=function(e){return o.p+"chunk."+({0:"HelloWorld-property-pane"}[e]||e)+"_"+{0:"d74bf9b629526c7e5db1"}[e]+".js"}(e),0!==u.src.indexOf(window.location.origin+"/")&&(u.crossOrigin="anonymous");var c=new Error;a=function(t){u.onerror=u.onload=null,clearTimeout(s);var n=r[e];if(0!==n){if(n){var o=t&&("load"===t.type?"missing":t.type),i=t&&t.target&&t.target.src;c.message="Loading chunk "+e+" failed.\n("+o+": "+i+")",c.name="ChunkLoadError",c.type=o,c.request=i,n[1](c)}r[e]=void 0}};var s=setTimeout(function(){a({type:"timeout",target:u})},12e4);u.onerror=u.onload=a,document.head.appendChild(u)}return Promise.all(t)},o.m=e,o.c=n,o.d=function(e,t,n){o.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:n})},o.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},o.t=function(e,t){if(1&t&&(e=o(e)),8&t)return e;if(4&t&&"object"==typeof e&&e&&e.__esModule)return e;var n=Object.create(null);if(o.r(n),Object.defineProperty(n,"default",{enumerable:!0,value:e}),2&t&&"string"!=typeof e)for(var r in e)o.d(n,r,function(t){return e[t]}.bind(null,r));return n},o.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return o.d(t,"a",t),t},o.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},o.p="",o.oe=function(e){throw console.error(e),e};var i=window.webpackJsonp_717e8239_32cd_4afc_af51_216d2d1dfc41_0_0_1=window.webpackJsonp_717e8239_32cd_4afc_af51_216d2d1dfc41_0_0_1||[],a=i.push.bind(i);i.push=t,i=i.slice();for(var u=0;u<i.length;u++)t(i[u]);var c=a;return function(){var e,t=document.getElementsByTagName("script"),n=/hello-world-adaptive-card-extension_65bae68925aa1aeff998\.js/i;if(t&&t.length)for(var r=0;r<t.length;r++)if(t[r]){var i=t[r].getAttribute("src");if(i&&i.match(n)){e=i.substring(0,i.lastIndexOf("/")+1);break}}if(!e)for(var a in window.__setWebpackPublicPathLoaderSrcRegistry__)if(a&&a.match(n)){e=a.substring(0,a.lastIndexOf("/")+1);break}o.p=e}(),o(o.s="T4Pj")}({"5nUA":function(e,t,n){e.exports=n.p+"SharePointLogo_080ce1f0d32aa6185206d1b09cf37de9.svg"},T4Pj:function(e,t,n){"use strict";n.r(t),n.d(t,"QUICK_VIEW_REGISTRY_ID",function(){return f});var r,o=n("lz/E"),i=(r=function(e,t){return(r=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n])})(e,t)},function(e,t){function n(){this.constructor=e}r(e,t),e.prototype=null===t?Object.create(t):(n.prototype=t.prototype,new n)}),a=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return i(t,e),Object.defineProperty(t.prototype,"cardButtons",{get:function(){return[{title:"Preview Messages",action:{type:"QuickView",parameters:{view:f}}},{title:"Outlook",action:{type:"ExternalLink",parameters:{target:"http://outlook.office.com"}}}]},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"data",{get:function(){return{primaryText:"You have "+this.state.unreadCount+" unread messages",unreadCount:this.state.unreadCount}},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"onCardSelection",{get:function(){return{type:"QuickView",parameters:{view:f}}},enumerable:!0,configurable:!0}),t}(o.BaseBasicCardView),u=function(){var e=function(t,n){return(e=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n])})(t,n)};return function(t,n){function r(){this.constructor=t}e(t,n),t.prototype=null===n?Object.create(n):(r.prototype=n.prototype,new r)}}(),c=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return u(t,e),Object.defineProperty(t.prototype,"data",{get:function(){return{emails:this.state.emails}},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"template",{get:function(){return n("YBiG")},enumerable:!0,configurable:!0}),t}(o.BaseAdaptiveCardView),s=function(){var e=function(t,n){return(e=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n])})(t,n)};return function(t,n){function r(){this.constructor=t}e(t,n),t.prototype=null===n?Object.create(n):(r.prototype=n.prototype,new r)}}(),l=function(e,t,n,r){return new(n||(n=Promise))(function(o,i){function a(e){try{c(r.next(e))}catch(e){i(e)}}function u(e){try{c(r.throw(e))}catch(e){i(e)}}function c(e){var t;e.done?o(e.value):(t=e.value,t instanceof n?t:new n(function(e){e(t)})).then(a,u)}c((r=r.apply(e,t||[])).next())})},p=function(e,t){var n,r,o,i,a={label:0,sent:function(){if(1&o[0])throw o[1];return o[1]},trys:[],ops:[]};return i={next:u(0),throw:u(1),return:u(2)},"function"==typeof Symbol&&(i[Symbol.iterator]=function(){return this}),i;function u(i){return function(u){return function(i){if(n)throw new TypeError("Generator is already executing.");for(;a;)try{if(n=1,r&&(o=2&i[0]?r.return:i[0]?r.throw||((o=r.return)&&o.call(r),0):r.next)&&!(o=o.call(r,i[1])).done)return o;switch(r=0,o&&(i=[2&i[0],o.value]),i[0]){case 0:case 1:o=i;break;case 4:return a.label++,{value:i[1],done:!1};case 5:a.label++,r=i[1],i=[0];continue;case 7:i=a.ops.pop(),a.trys.pop();continue;default:if(!((o=(o=a.trys).length>0&&o[o.length-1])||6!==i[0]&&2!==i[0])){a=0;continue}if(3===i[0]&&(!o||i[1]>o[0]&&i[1]<o[3])){a.label=i[1];break}if(6===i[0]&&a.label<o[1]){a.label=o[1],o=i;break}if(o&&a.label<o[2]){a.label=o[2],a.ops.push(i);break}o[2]&&a.ops.pop(),a.trys.pop();continue}i=t.call(e,a)}catch(e){i=[6,e],r=0}finally{n=o=0}if(5&i[0])throw i[1];return{value:i[0]?i[1]:void 0,done:!0}}([i,u])}}},f="HelloWorld_QUICK_VIEW",d=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return s(t,e),t.prototype.onInit=function(){return l(this,void 0,void 0,function(){return p(this,function(e){switch(e.label){case 0:return this.state={description:this.properties.description,unreadCount:0,emails:[]},[4,this.getUnreadCount()];case 1:return e.sent(),[4,this.getEmailDetails()];case 2:return e.sent(),this.cardNavigator.register("HelloWorld_CARD_VIEW",function(){return new a}),this.quickViewNavigator.register(f,function(){return new c}),[2,Promise.resolve()]}})})},t.prototype.getUnreadCount=function(){return l(this,void 0,void 0,function(){var e,t,n=this;return p(this,function(r){switch(r.label){case 0:return[4,this.context.msGraphClientFactory.getClient()];case 1:e=r.sent(),r.label=2;case 2:return r.trys.push([2,4,,5]),[4,e.api("/me/messages").version("v1.0").filter("isRead ne true&$count=true&$top=999").get(function(e,t,r){n.setState({unreadCount:t.value.length}),console.log("getUnreadCount RESPONSE",t)})];case 3:return r.sent(),[3,5];case 4:return t=r.sent(),console.log(t),[3,5];case 5:return[2]}})})},t.prototype.getEmailDetails=function(){return l(this,void 0,void 0,function(){var e,t,n=this;return p(this,function(r){switch(r.label){case 0:return[4,this.context.msGraphClientFactory.getClient()];case 1:e=r.sent(),r.label=2;case 2:return r.trys.push([2,4,,5]),[4,e.api("/me/messages").version("v1.0").filter("isRead ne true&$count=true&$top=999").get(function(e,t,r){t.value.forEach(function(e){n.state.emails.push({webLink:e.webLink,subject:e.subject,sender:e.sender.emailAddress.name})})})];case 3:return r.sent(),console.log(this.state.emails),[3,5];case 4:return t=r.sent(),console.log(t),[3,5];case 5:return[2]}})})},Object.defineProperty(t.prototype,"title",{get:function(){return this.properties.title},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"iconProperty",{get:function(){return this.properties.iconProperty||n("5nUA")},enumerable:!0,configurable:!0}),t.prototype.loadPropertyPaneResources=function(){var e=this;return n.e(0).then(n.bind(null,"09z0")).then(function(t){e._deferredPropertyPane=new t.HelloWorldPropertyPane})},t.prototype.renderCard=function(){return"HelloWorld_CARD_VIEW"},t.prototype.getPropertyPaneConfiguration=function(){return this._deferredPropertyPane.getPropertyPaneConfiguration()},t}(o.BaseAdaptiveCardExtension);t.default=d},YBiG:function(e){e.exports=JSON.parse('{"schema":"http://adaptivecards.io/schemas/adaptive-card.json","type":"AdaptiveCard","version":"1.2","body":[{"type":"Container","$data":"${emails}","items":[{"type":"TextBlock","text":"${sender}:","weight":"bolder","size":"medium"},{"type":"TextBlock","text":"${subject}","wrap":true}],"selectAction":{"type":"Action.OpenUrl","url":"${webLink}"}}]}')},"lz/E":function(t,n){t.exports=e}})});