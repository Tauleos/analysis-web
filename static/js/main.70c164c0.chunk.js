(window.webpackJsonp=window.webpackJsonp||[]).push([[0],{128:function(e,a,t){e.exports=t.p+"static/media/logo.5d5d9eef.svg"},141:function(e,a,t){e.exports=t(314)},146:function(e,a,t){},157:function(e,a,t){},314:function(e,a,t){"use strict";t.r(a);var n=t(1),o=t.n(n),r=t(6),c=t.n(r),l=(t(146),t(313),t(136)),s=(t(150),t(8)),i=(t(84),t(19)),d=(t(153),t(137)),m=(t(155),t(43)),u=t(125),p=t(126),f=t(138),h=t(127),g=t(139),v=t(128),w=t.n(v),y=(t(157),t(129)),E=t.n(y).a.create({baseURL:""}),b=function(e){function a(){var e,t;Object(u.a)(this,a);for(var n=arguments.length,r=new Array(n),c=0;c<n;c++)r[c]=arguments[c];return(t=Object(f.a)(this,(e=Object(h.a)(a)).call.apply(e,[this].concat(r)))).state={filename:"",targetUrl:void 0},t.error_message="\u4e0a\u4f20\u5931\u8d25\uff0c\u8bf7\u8054\u7cfb\u53ef\u601c\u7684\u8001\u516c\uff01",t.upload_props={name:"file",action:"".concat("","/upload"),onChange:function(e){var a=e.file;"done"===a.status?a.response.success?(t.setState({filename:a.response.filename}),m.a.success("\u4e0a\u4f20\u6210\u529f\uff0c\u8bf7\u70b9\u51fb\u5904\u7406")):m.a.error(t.error_message):"error"===a.status&&m.a.error(t.error_message)}},t.upload=function(){E.post("/exe",{filename:t.state.filename}).then(function(e){t.setState({targetUrl:e.data}),d.a.info({title:"\u5904\u7406\u5b8c\u6210\uff0c\u8bf7\u70b9\u51fb\u4e0b\u65b9\u6309\u94ae\u4e0b\u8f7d",content:o.a.createElement("div",null,o.a.createElement(i.a,{type:"primary",icon:"cloud-download",onClick:t.downloadUrl}))})}).catch(function(e){m.a.error(t.error_message)})},t.downloadUrl=function(){var e=document.createElement("iframe");e.style.display="none",e.src="".concat("").concat(t.state.targetUrl),document.body.appendChild(e),setTimeout(function(){document.body.removeChild(e)},100)},t}return Object(g.a)(a,e),Object(p.a)(a,[{key:"render",value:function(){return o.a.createElement("div",{className:"App"},o.a.createElement("header",{className:"App-header"},o.a.createElement("img",{src:w.a,className:"App-logo",alt:"logo"})),o.a.createElement("div",{className:"content"},o.a.createElement(l.a,Object.assign({},this.upload_props,{className:"upload"}),o.a.createElement(i.a,null,o.a.createElement(s.a,{type:"upload"}),"\u8bf7\u70b9\u51fb\u4e0a\u4f20Excel")),o.a.createElement(i.a,{type:"primary",disabled:!this.state.filename,onClick:this.upload},"\u70b9\u51fb\u5904\u7406")))}}]),a}(n.Component);Boolean("localhost"===window.location.hostname||"[::1]"===window.location.hostname||window.location.hostname.match(/^127(?:\.(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)){3}$/));c.a.render(o.a.createElement(b,null),document.getElementById("root")),"serviceWorker"in navigator&&navigator.serviceWorker.ready.then(function(e){e.unregister()})}},[[141,1,2]]]);
//# sourceMappingURL=main.70c164c0.chunk.js.map