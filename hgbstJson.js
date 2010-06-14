if (!this.JSON) { JSON = function () { function f(n) { return n < 10 ? '0' + n : n; } Date.prototype.toJSON = function () { return this.getUTCFullYear() + '-' + f(this.getUTCMonth() + 1) + '-' + f(this.getUTCDate()) + 'T' + f(this.getUTCHours()) + ':' + f(this.getUTCMinutes()) + ':' + f(this.getUTCSeconds()) + 'Z'; }; var m = { '\b': '\\b', '\t': '\\t', '\n': '\\n', '\f': '\\f', '\r': '\\r', '"' : '\\"', '\\': '\\\\' }; function stringify(value, whitelist) { var a, i, k, l, r = /["\\\x00-\x1f\x7f-\x9f]/g, v; switch (typeof value) { case 'string': return r.test(value) ? '"' + value.replace(r, function (a) { var c = m[a]; if (c) { return c; } c = a.charCodeAt(); return '\\u00' + Math.floor(c / 16).toString(16) + (c % 16).toString(16); }) + '"' : '"' + value + '"'; case 'number': return isFinite(value) ? String(value) : 'null'; case 'boolean': case 'null': return String(value); case 'object': if (!value) { return 'null'; } if (typeof value.toJSON === 'function') { return stringify(value.toJSON()); } a = []; if (typeof value.length === 'number' && !(value.propertyIsEnumerable('length'))) { l = value.length; for (i = 0; i < l; i += 1) { a.push(stringify(value[i], whitelist) || 'null'); } return '[' + a.join(',') + ']'; } if (whitelist) { l = whitelist.length; for (i = 0; i < l; i += 1) { k = whitelist[i]; if (typeof k === 'string') { v = stringify(value[k], whitelist); if (v) { a.push(stringify(k) + ':' + v); } } } } else { for (k in value) { if (typeof k === 'string') { v = stringify(value[k], whitelist); if (v) { a.push(stringify(k) + ':' + v); } } } } return '{' + a.join(',') + '}'; } } return { stringify: stringify, parse: function (text, filter) { var j; function walk(k, v) { var i, n; if (v && typeof v === 'object') { for (i in v) { if (Object.prototype.hasOwnProperty.apply(v, [i])) { n = walk(i, v[i]); if (n !== undefined) { v[i] = n; } } } } return filter(k, v); } if (/^[\],:{}\s]*$/.test(text.replace(/\\./g, '@'). replace(/"[^"\\\n\r]*"|true|false|null|-?\d+(?:\.\d*)?(:?[eE][+\-]?\d+)?/g, ']'). replace(/(?:^|:|,)(?:\s*\[)+/g, ''))) { j = eval('(' + text + ')'); return typeof filter === 'function' ? walk('', j) : j; } throw new SyntaxError('parseJSON'); } }; }(); }

var FCApp = new ActiveXObject("FCFLCompat.FCApplication");
FCApp.WorkingDirectory = 'C:\\apps\\agent\\pages\\';
FCApp.Initialize();

var FCSession = FCApp.CreateSession();
FCSession.Login('sa', 'sa', 'user');
FCSession.ThrowErrors = false;

var maximumDepth = 3;
var listTitle = "CR_DESC";

var jsonList = {};
jsonList.title = listTitle;
jsonList.Children = addChildren();

WScript.Echo(JSON.stringify(jsonList));

function addChildren() {
   var args = arguments;
   var len = args.length;
   var list = getList.apply(this, args);

   if(list.RecordCount > 0) {
      var currentLevel = [];
      for (var x = 0; x < list.RecordCount; x++) {
         var currentLevel_item = {};
         var thisElement = list("title") + "";
         currentLevel_item.title = thisElement;

         if(len == 0) currentLevel_item.Children = addChildren(thisElement);
         else if(len == 1) currentLevel_item.Children = addChildren(args[0], thisElement);
         else if(len == 2) currentLevel_item.Children = addChildren(args[0], args[1], thisElement);
         else if(len == 3) currentLevel_item.Children = addChildren(args[0], args[1], args[2], thisElement);
         else if(len == 4) currentLevel_item.Children = addChildren(args[0], args[1], args[2], args[3], thisElement);
         else if(len == 4) currentLevel_item.Children = addChildren(args[0], args[1], args[2], args[3], args[4], thisElement);

         currentLevel.push(currentLevel_item);

         list.MoveNext();
      }

      return currentLevel;
   }
}

function getList() {
   var args = arguments;
   var len = args.length;
   var list = {};cd .

   if(len > maximumDepth-1) list.Recordcount = 0;
   else if(len == 0) list = FCApp.GetHgbstList(listTitle);
   else if(len == 1) list = FCApp.GetHgbstList(listTitle, args[0]);
   else if(len == 2) list = FCApp.GetHgbstList(listTitle, args[0], args[1]);
   else if(len == 3) list = FCApp.GetHgbstList(listTitle, args[0], args[1], args[2]);
   else if(len == 4) list = FCApp.GetHgbstList(listTitle, args[0], args[1], args[2], args[3]);
   else if(len == 5) list = FCApp.GetHgbstList(listTitle, args[0], args[1], args[2], args[3], args[4]);

   return list;
}