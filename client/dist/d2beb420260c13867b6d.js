function _typeof(obj) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (obj) { return typeof obj; } : function (obj) { return obj && "function" == typeof Symbol && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; }, _typeof(obj); }
function _readOnlyError(name) { throw new TypeError("\"" + name + "\" is read-only"); }
function _regeneratorRuntime() { "use strict"; /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/facebook/regenerator/blob/main/LICENSE */ _regeneratorRuntime = function _regeneratorRuntime() { return exports; }; var exports = {}, Op = Object.prototype, hasOwn = Op.hasOwnProperty, defineProperty = Object.defineProperty || function (obj, key, desc) { obj[key] = desc.value; }, $Symbol = "function" == typeof Symbol ? Symbol : {}, iteratorSymbol = $Symbol.iterator || "@@iterator", asyncIteratorSymbol = $Symbol.asyncIterator || "@@asyncIterator", toStringTagSymbol = $Symbol.toStringTag || "@@toStringTag"; function define(obj, key, value) { return Object.defineProperty(obj, key, { value: value, enumerable: !0, configurable: !0, writable: !0 }), obj[key]; } try { define({}, ""); } catch (err) { define = function define(obj, key, value) { return obj[key] = value; }; } function wrap(innerFn, outerFn, self, tryLocsList) { var protoGenerator = outerFn && outerFn.prototype instanceof Generator ? outerFn : Generator, generator = Object.create(protoGenerator.prototype), context = new Context(tryLocsList || []); return defineProperty(generator, "_invoke", { value: makeInvokeMethod(innerFn, self, context) }), generator; } function tryCatch(fn, obj, arg) { try { return { type: "normal", arg: fn.call(obj, arg) }; } catch (err) { return { type: "throw", arg: err }; } } exports.wrap = wrap; var ContinueSentinel = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} var IteratorPrototype = {}; define(IteratorPrototype, iteratorSymbol, function () { return this; }); var getProto = Object.getPrototypeOf, NativeIteratorPrototype = getProto && getProto(getProto(values([]))); NativeIteratorPrototype && NativeIteratorPrototype !== Op && hasOwn.call(NativeIteratorPrototype, iteratorSymbol) && (IteratorPrototype = NativeIteratorPrototype); var Gp = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(IteratorPrototype); function defineIteratorMethods(prototype) { ["next", "throw", "return"].forEach(function (method) { define(prototype, method, function (arg) { return this._invoke(method, arg); }); }); } function AsyncIterator(generator, PromiseImpl) { function invoke(method, arg, resolve, reject) { var record = tryCatch(generator[method], generator, arg); if ("throw" !== record.type) { var result = record.arg, value = result.value; return value && "object" == _typeof(value) && hasOwn.call(value, "__await") ? PromiseImpl.resolve(value.__await).then(function (value) { invoke("next", value, resolve, reject); }, function (err) { invoke("throw", err, resolve, reject); }) : PromiseImpl.resolve(value).then(function (unwrapped) { result.value = unwrapped, resolve(result); }, function (error) { return invoke("throw", error, resolve, reject); }); } reject(record.arg); } var previousPromise; defineProperty(this, "_invoke", { value: function value(method, arg) { function callInvokeWithMethodAndArg() { return new PromiseImpl(function (resolve, reject) { invoke(method, arg, resolve, reject); }); } return previousPromise = previousPromise ? previousPromise.then(callInvokeWithMethodAndArg, callInvokeWithMethodAndArg) : callInvokeWithMethodAndArg(); } }); } function makeInvokeMethod(innerFn, self, context) { var state = "suspendedStart"; return function (method, arg) { if ("executing" === state) throw new Error("Generator is already running"); if ("completed" === state) { if ("throw" === method) throw arg; return doneResult(); } for (context.method = method, context.arg = arg;;) { var delegate = context.delegate; if (delegate) { var delegateResult = maybeInvokeDelegate(delegate, context); if (delegateResult) { if (delegateResult === ContinueSentinel) continue; return delegateResult; } } if ("next" === context.method) context.sent = context._sent = context.arg;else if ("throw" === context.method) { if ("suspendedStart" === state) throw state = "completed", context.arg; context.dispatchException(context.arg); } else "return" === context.method && context.abrupt("return", context.arg); state = "executing"; var record = tryCatch(innerFn, self, context); if ("normal" === record.type) { if (state = context.done ? "completed" : "suspendedYield", record.arg === ContinueSentinel) continue; return { value: record.arg, done: context.done }; } "throw" === record.type && (state = "completed", context.method = "throw", context.arg = record.arg); } }; } function maybeInvokeDelegate(delegate, context) { var methodName = context.method, method = delegate.iterator[methodName]; if (undefined === method) return context.delegate = null, "throw" === methodName && delegate.iterator.return && (context.method = "return", context.arg = undefined, maybeInvokeDelegate(delegate, context), "throw" === context.method) || "return" !== methodName && (context.method = "throw", context.arg = new TypeError("The iterator does not provide a '" + methodName + "' method")), ContinueSentinel; var record = tryCatch(method, delegate.iterator, context.arg); if ("throw" === record.type) return context.method = "throw", context.arg = record.arg, context.delegate = null, ContinueSentinel; var info = record.arg; return info ? info.done ? (context[delegate.resultName] = info.value, context.next = delegate.nextLoc, "return" !== context.method && (context.method = "next", context.arg = undefined), context.delegate = null, ContinueSentinel) : info : (context.method = "throw", context.arg = new TypeError("iterator result is not an object"), context.delegate = null, ContinueSentinel); } function pushTryEntry(locs) { var entry = { tryLoc: locs[0] }; 1 in locs && (entry.catchLoc = locs[1]), 2 in locs && (entry.finallyLoc = locs[2], entry.afterLoc = locs[3]), this.tryEntries.push(entry); } function resetTryEntry(entry) { var record = entry.completion || {}; record.type = "normal", delete record.arg, entry.completion = record; } function Context(tryLocsList) { this.tryEntries = [{ tryLoc: "root" }], tryLocsList.forEach(pushTryEntry, this), this.reset(!0); } function values(iterable) { if (iterable) { var iteratorMethod = iterable[iteratorSymbol]; if (iteratorMethod) return iteratorMethod.call(iterable); if ("function" == typeof iterable.next) return iterable; if (!isNaN(iterable.length)) { var i = -1, next = function next() { for (; ++i < iterable.length;) if (hasOwn.call(iterable, i)) return next.value = iterable[i], next.done = !1, next; return next.value = undefined, next.done = !0, next; }; return next.next = next; } } return { next: doneResult }; } function doneResult() { return { value: undefined, done: !0 }; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, defineProperty(Gp, "constructor", { value: GeneratorFunctionPrototype, configurable: !0 }), defineProperty(GeneratorFunctionPrototype, "constructor", { value: GeneratorFunction, configurable: !0 }), GeneratorFunction.displayName = define(GeneratorFunctionPrototype, toStringTagSymbol, "GeneratorFunction"), exports.isGeneratorFunction = function (genFun) { var ctor = "function" == typeof genFun && genFun.constructor; return !!ctor && (ctor === GeneratorFunction || "GeneratorFunction" === (ctor.displayName || ctor.name)); }, exports.mark = function (genFun) { return Object.setPrototypeOf ? Object.setPrototypeOf(genFun, GeneratorFunctionPrototype) : (genFun.__proto__ = GeneratorFunctionPrototype, define(genFun, toStringTagSymbol, "GeneratorFunction")), genFun.prototype = Object.create(Gp), genFun; }, exports.awrap = function (arg) { return { __await: arg }; }, defineIteratorMethods(AsyncIterator.prototype), define(AsyncIterator.prototype, asyncIteratorSymbol, function () { return this; }), exports.AsyncIterator = AsyncIterator, exports.async = function (innerFn, outerFn, self, tryLocsList, PromiseImpl) { void 0 === PromiseImpl && (PromiseImpl = Promise); var iter = new AsyncIterator(wrap(innerFn, outerFn, self, tryLocsList), PromiseImpl); return exports.isGeneratorFunction(outerFn) ? iter : iter.next().then(function (result) { return result.done ? result.value : iter.next(); }); }, defineIteratorMethods(Gp), define(Gp, toStringTagSymbol, "Generator"), define(Gp, iteratorSymbol, function () { return this; }), define(Gp, "toString", function () { return "[object Generator]"; }), exports.keys = function (val) { var object = Object(val), keys = []; for (var key in object) keys.push(key); return keys.reverse(), function next() { for (; keys.length;) { var key = keys.pop(); if (key in object) return next.value = key, next.done = !1, next; } return next.done = !0, next; }; }, exports.values = values, Context.prototype = { constructor: Context, reset: function reset(skipTempReset) { if (this.prev = 0, this.next = 0, this.sent = this._sent = undefined, this.done = !1, this.delegate = null, this.method = "next", this.arg = undefined, this.tryEntries.forEach(resetTryEntry), !skipTempReset) for (var name in this) "t" === name.charAt(0) && hasOwn.call(this, name) && !isNaN(+name.slice(1)) && (this[name] = undefined); }, stop: function stop() { this.done = !0; var rootRecord = this.tryEntries[0].completion; if ("throw" === rootRecord.type) throw rootRecord.arg; return this.rval; }, dispatchException: function dispatchException(exception) { if (this.done) throw exception; var context = this; function handle(loc, caught) { return record.type = "throw", record.arg = exception, context.next = loc, caught && (context.method = "next", context.arg = undefined), !!caught; } for (var i = this.tryEntries.length - 1; i >= 0; --i) { var entry = this.tryEntries[i], record = entry.completion; if ("root" === entry.tryLoc) return handle("end"); if (entry.tryLoc <= this.prev) { var hasCatch = hasOwn.call(entry, "catchLoc"), hasFinally = hasOwn.call(entry, "finallyLoc"); if (hasCatch && hasFinally) { if (this.prev < entry.catchLoc) return handle(entry.catchLoc, !0); if (this.prev < entry.finallyLoc) return handle(entry.finallyLoc); } else if (hasCatch) { if (this.prev < entry.catchLoc) return handle(entry.catchLoc, !0); } else { if (!hasFinally) throw new Error("try statement without catch or finally"); if (this.prev < entry.finallyLoc) return handle(entry.finallyLoc); } } } }, abrupt: function abrupt(type, arg) { for (var i = this.tryEntries.length - 1; i >= 0; --i) { var entry = this.tryEntries[i]; if (entry.tryLoc <= this.prev && hasOwn.call(entry, "finallyLoc") && this.prev < entry.finallyLoc) { var finallyEntry = entry; break; } } finallyEntry && ("break" === type || "continue" === type) && finallyEntry.tryLoc <= arg && arg <= finallyEntry.finallyLoc && (finallyEntry = null); var record = finallyEntry ? finallyEntry.completion : {}; return record.type = type, record.arg = arg, finallyEntry ? (this.method = "next", this.next = finallyEntry.finallyLoc, ContinueSentinel) : this.complete(record); }, complete: function complete(record, afterLoc) { if ("throw" === record.type) throw record.arg; return "break" === record.type || "continue" === record.type ? this.next = record.arg : "return" === record.type ? (this.rval = this.arg = record.arg, this.method = "return", this.next = "end") : "normal" === record.type && afterLoc && (this.next = afterLoc), ContinueSentinel; }, finish: function finish(finallyLoc) { for (var i = this.tryEntries.length - 1; i >= 0; --i) { var entry = this.tryEntries[i]; if (entry.finallyLoc === finallyLoc) return this.complete(entry.completion, entry.afterLoc), resetTryEntry(entry), ContinueSentinel; } }, catch: function _catch(tryLoc) { for (var i = this.tryEntries.length - 1; i >= 0; --i) { var entry = this.tryEntries[i]; if (entry.tryLoc === tryLoc) { var record = entry.completion; if ("throw" === record.type) { var thrown = record.arg; resetTryEntry(entry); } return thrown; } } throw new Error("illegal catch attempt"); }, delegateYield: function delegateYield(iterable, resultName, nextLoc) { return this.delegate = { iterator: values(iterable), resultName: resultName, nextLoc: nextLoc }, "next" === this.method && (this.arg = undefined), ContinueSentinel; } }, exports; }
function asyncGeneratorStep(gen, resolve, reject, _next, _throw, key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { Promise.resolve(value).then(_next, _throw); } }
function _asyncToGenerator(fn) { return function () { var self = this, args = arguments; return new Promise(function (resolve, reject) { var gen = fn.apply(self, args); function _next(value) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "next", value); } function _throw(err) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "throw", err); } _next(undefined); }); }; }
/*
 * Copyright (c) Blake Woodford. All rights reserved. Licensed under the MIT license.
 */

// #region Imports
import { getUserProfile } from "../helpers/sso-helper";
import { filterUserProfileInfo } from "./../helpers/documentHelper";
var auth = null; // Auth token storage
var verifiedEmail = null; // Signup Email storage
var currentMSEmail = null; // Microsoft Email storage
//#endregion

// #region Initialize Office
Office.onReady(function (info) {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
    // Check if Excel version is supported
    console.log("Sorry, this add-in only works with newer versions of Excel.");
  }
  if (info.host === Office.HostType.Excel) {
    getUserProfile(checkMSVerification); // Check if user is verified on startup
  }
});
// #endregion

//#region User Verifications

function updateEmail(id, email) {
  // TODO: Update the email address in the settings of the task pane
  var textarea = document.getElementById(id);
  textarea.placeholder = email;
  verifiedEmail = email; //updates global value
  var button = document.getElementById("auth-button");
  button.style.backgroundColor = "#a1a1a1";
  button.style.cursor = "default";
  console.log("Email updated.");
  return;
}

// Sends verification packet to server to verify on Microsoft Email || Input: Microsoft Email || Output: JSON Verification Packet
function sendMSVerifyPacket(_x) {
  return _sendMSVerifyPacket.apply(this, arguments);
} // Checks verification of Micrsoft Email || Verified: logs auth, Not-Verified: does nothing
function _sendMSVerifyPacket() {
  _sendMSVerifyPacket = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee(msEmail) {
    var response, data;
    return _regeneratorRuntime().wrap(function _callee$(_context) {
      while (1) switch (_context.prev = _context.next) {
        case 0:
          _context.prev = 0;
          _context.next = 3;
          return fetch("https://testing-f03s.onrender.com/verify-email", {
            method: "POST",
            headers: {
              "Content-Type": "application/json"
            },
            body: JSON.stringify({
              email: msEmail
            })
          });
        case 3:
          response = _context.sent;
          if (!response.ok) {
            _context.next = 11;
            break;
          }
          _context.next = 7;
          return response.json();
        case 7:
          data = _context.sent;
          return _context.abrupt("return", data);
        case 11:
          console.error("Error:", response.status);
          return _context.abrupt("return", null);
        case 13:
          _context.next = 19;
          break;
        case 15:
          _context.prev = 15;
          _context.t0 = _context["catch"](0);
          console.error("Error:", _context.t0);
          return _context.abrupt("return", null);
        case 19:
        case "end":
          return _context.stop();
      }
    }, _callee, null, [[0, 15]]);
  }));
  return _sendMSVerifyPacket.apply(this, arguments);
}
function checkMSVerification(_x2) {
  return _checkMSVerification.apply(this, arguments);
}
function _checkMSVerification() {
  _checkMSVerification = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee2(result) {
    var msEmail, verificationStatus;
    return _regeneratorRuntime().wrap(function _callee2$(_context2) {
      while (1) switch (_context2.prev = _context2.next) {
        case 0:
          msEmail = filterUserProfileInfo(result)[0];
          if (!(msEmail === null)) {
            _context2.next = 6;
            break;
          }
          console.log("No Microsoft Email");
          return _context2.abrupt("return");
        case 6:
          currentMSEmail = msEmail;
        case 7:
          _context2.next = 9;
          return sendMSVerifyPacket(msEmail);
        case 9:
          verificationStatus = _context2.sent;
          if (verificationStatus) {
            _context2.next = 12;
            break;
          }
          return _context2.abrupt("return");
        case 12:
          if (!verificationStatus.verification) {
            _context2.next = 17;
            break;
          }
          // On success, update authtoken and stored email
          auth = verificationStatus.authToken;
          updateEmail("signup-email", verificationStatus.stripeEmail);
          _context2.next = 19;
          break;
        case 17:
          // TODO: Handle unverified user
          console.log("User not verified on MS Email");
          return _context2.abrupt("return");
        case 19:
        case "end":
          return _context2.stop();
      }
    }, _callee2);
  }));
  return _checkMSVerification.apply(this, arguments);
}
function verifyUser(_x3) {
  return _verifyUser.apply(this, arguments);
}
function _verifyUser() {
  _verifyUser = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee3(signUpEmail) {
    var response, data;
    return _regeneratorRuntime().wrap(function _callee3$(_context3) {
      while (1) switch (_context3.prev = _context3.next) {
        case 0:
          _context3.prev = 0;
          _context3.next = 3;
          return fetch("https://testing-f03s.onrender.com/verify-user", {
            method: "POST",
            headers: {
              "Content-Type": "application/json"
            },
            body: JSON.stringify({
              email: signUpEmail,
              msEmail: currentMSEmail
            })
          });
        case 3:
          response = _context3.sent;
          if (!response.ok) {
            _context3.next = 11;
            break;
          }
          _context3.next = 7;
          return response.json();
        case 7:
          data = _context3.sent;
          return _context3.abrupt("return", data);
        case 11:
          console.error("Error:", response.status);
        case 12:
          _context3.next = 17;
          break;
        case 14:
          _context3.prev = 14;
          _context3.t0 = _context3["catch"](0);
          console.error("Error:", _context3.t0);
        case 17:
        case "end":
          return _context3.stop();
      }
    }, _callee3, null, [[0, 14]]);
  }));
  return _verifyUser.apply(this, arguments);
}
function authenticateUser(_x4, _x5) {
  return _authenticateUser.apply(this, arguments);
}
function _authenticateUser() {
  _authenticateUser = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee4(emailId, statusId) {
    var signUpEmail, status, verificationPacket;
    return _regeneratorRuntime().wrap(function _callee4$(_context4) {
      while (1) switch (_context4.prev = _context4.next) {
        case 0:
          if (verifiedEmail === null) {
            _context4.next = 3;
            break;
          }
          console.log("User already authenticated");
          return _context4.abrupt("return");
        case 3:
          signUpEmail = document.getElementById(emailId).value;
          status = document.getElementById(statusId);
          if (!(signUpEmail === "")) {
            _context4.next = 8;
            break;
          }
          status.value = "Please enter an email address.";
          return _context4.abrupt("return");
        case 8:
          _context4.prev = 8;
          _context4.next = 11;
          return verifyUser(signUpEmail);
        case 11:
          verificationPacket = _context4.sent;
          if (!verificationPacket.verification) {
            _context4.next = 19;
            break;
          }
          // On success, update authtoken and stored email
          auth = verificationPacket.authToken;
          updateEmail(emailId, signUpEmail);
          status.value = verificationPacket.verificationMessage;
          return _context4.abrupt("return");
        case 19:
          status.value = verificationPacket.verificationMessage;
          return _context4.abrupt("return");
        case 21:
          _context4.next = 28;
          break;
        case 23:
          _context4.prev = 23;
          _context4.t0 = _context4["catch"](8);
          console.error("Error:", _context4.t0);
          status.value = "Unable to verify email, please try again.";
          return _context4.abrupt("return");
        case 28:
        case "end":
          return _context4.stop();
      }
    }, _callee4, null, [[8, 23]]);
  }));
  return _authenticateUser.apply(this, arguments);
}
function sendDeAuth(_x6) {
  return _sendDeAuth.apply(this, arguments);
}
function _sendDeAuth() {
  _sendDeAuth = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee5(emailId) {
    var response, data;
    return _regeneratorRuntime().wrap(function _callee5$(_context5) {
      while (1) switch (_context5.prev = _context5.next) {
        case 0:
          if (!(verifiedEmail === null || currentMSEmail === undefined)) {
            _context5.next = 3;
            break;
          }
          console.log("No current user");
          return _context5.abrupt("return", {
            message: "No current user"
          });
        case 3:
          _context5.prev = 3;
          _context5.next = 6;
          return fetch("https://testing-f03s.onrender.com/deauth-user", {
            method: "POST",
            headers: {
              "Content-Type": "application/json"
            },
            body: JSON.stringify({
              email: verifiedEmail,
              msEmail: currentMSEmail,
              authToken: auth
            })
          });
        case 6:
          response = _context5.sent;
          if (!response.ok) {
            _context5.next = 18;
            break;
          }
          _context5.next = 10;
          return response.json();
        case 10:
          data = _context5.sent;
          auth = null;
          verifiedEmail = null;
          document.getElementById(emailId).value = "";
          document.getElementById(emailId).placeholder = "Email";
          return _context5.abrupt("return", data);
        case 18:
          console.error("Error:", response.status);
          return _context5.abrupt("return", {
            message: "Unable to deauthenticate user, please try again."
          });
        case 20:
          _context5.next = 26;
          break;
        case 22:
          _context5.prev = 22;
          _context5.t0 = _context5["catch"](3);
          console.error("Error:", _context5.t0);
          return _context5.abrupt("return", {
            message: "Unable to deauthenticate user, please try again."
          });
        case 26:
        case "end":
          return _context5.stop();
      }
    }, _callee5, null, [[3, 22]]);
  }));
  return _sendDeAuth.apply(this, arguments);
}
function deAuthenticateUser(_x7, _x8) {
  return _deAuthenticateUser.apply(this, arguments);
} //#endregion
//#region Handle switching between tabs
function _deAuthenticateUser() {
  _deAuthenticateUser = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee6(emailId, statusId) {
    var deauthpacket, status;
    return _regeneratorRuntime().wrap(function _callee6$(_context6) {
      while (1) switch (_context6.prev = _context6.next) {
        case 0:
          _context6.next = 2;
          return sendDeAuth(emailId);
        case 2:
          deauthpacket = _context6.sent;
          status = document.getElementById(statusId);
          deauthpacket.message, _readOnlyError("status");
        case 5:
        case "end":
          return _context6.stop();
      }
    }, _callee6);
  }));
  return _deAuthenticateUser.apply(this, arguments);
}
document.addEventListener("DOMContentLoaded", function () {
  "use strict";

  // Store references to the tabs and tab contents
  var tabs = document.querySelectorAll(".tab");
  var tabContents = document.querySelectorAll(".tab-content");

  // Function to switch tabs
  function switchTab(tabId) {
    // Hide all tab contents and deactivate all tabs
    tabContents.forEach(function (tabContent) {
      tabContent.style.display = "none";
    });
    tabs.forEach(function (tab) {
      tab.classList.remove("active");
    });

    // Show the selected tab content and activate the corresponding tab
    var selectedTabContent = document.getElementById(tabId);
    var selectedTab = document.querySelector('a[data-tab="' + tabId + '"]');
    selectedTabContent.style.display = "flex";
    selectedTab.classList.add("active");
  }

  // Add click event listeners to the tabs
  tabs.forEach(function (tab) {
    tab.addEventListener("click", function () {
      var tabId = this.getAttribute("data-tab");
      switchTab(tabId);
    });
  });

  // Set the initial active tab
  switchTab("TabCreate");
});

//#endregion

//#region Functions to load and display text

// Loader element to display elipsis while waiting for GPT response
var loadInterval = null;
function loader(element) {
  element.textContent = "";
  loadInterval = setInterval(function () {
    // Update the text content of the loading indicator
    element.textContent += ". ";

    // If the loading indicator has reached three dots, reset it
    if (element.textContent === ". . . . ") {
      element.textContent = "";
    }
  }, 300);
}
function typeText(element, wrappedText, uniqueId, type) {
  var index = 0;
  var currentSpan = null;
  if (wrappedText === undefined || wrappedText === "") {
    return;
  }
  var interval = setInterval(function () {
    if (index < wrappedText.length) {
      var char = wrappedText.charAt(index);
      if (char === "<" && wrappedText.charAt(index + 1) !== "/") {
        var closingBracketIndex = wrappedText.indexOf(">", index);
        if (closingBracketIndex !== -1) {
          var spanTag = wrappedText.substring(index, closingBracketIndex + 1);
          element.innerHTML += spanTag;
          currentSpan = element.lastElementChild;
          index = closingBracketIndex;
        }
      } else if (char === "<" && wrappedText.charAt(index + 1) === "/") {
        var _closingBracketIndex = wrappedText.indexOf(">", index);
        if (_closingBracketIndex !== -1) {
          currentSpan = null;
          index = _closingBracketIndex;
        }
      } else if (currentSpan !== null) {
        currentSpan.innerHTML += char;
      } else {
        element.innerHTML += char;
      }
      index++;
    } else {
      clearInterval(interval);
    }
  }, 20);
  if (uniqueId !== undefined) {
    document.getElementById("insert-".concat(uniqueId)).addEventListener("click", function () {
      insertFormula("insert-".concat(uniqueId));
    });
    document.getElementById("retry-".concat(uniqueId)).addEventListener("click", function () {
      retryQuery(uniqueId, type);
    });
  }
}
function generateUniqueId() {
  var timestamp = Date.now();
  var randomNumber = Math.random();
  var hexadecimalString = randomNumber.toString(16);
  return "".concat(timestamp, "-").concat(hexadecimalString);
}
function outputText(isAi, value, uniqueId) {
  return "\n        <div class=\"wrapper ".concat(isAi && "ai", "\">\n            <div class=\"chat\">\n                <div class=\"profile\">").concat(isAi ? "🥷" : "👤", "</div>\n                <div class=\"message ").concat(isAi && "ai-message", "\" id=").concat(isAi ? "output-" : "input-").concat(uniqueId, ">").concat(value, "</div>\n                ").concat(isAi ? "<button id=\"insert-".concat(uniqueId, "\" title=\"Paste into Current Cell\" class=\"insert\"><i class=\"fas fa-clipboard\"></i></button>") : "<button id=\"retry-".concat(uniqueId, "\" title=\"Resend\" class=\"insert\"><i class=\"fas fa-redo\"></i></button>"), "\n            </div>\n        </div>\n        ");
}
//#endregion

//#region Functions to handle insert/retry button
function retryQuery(_x9, _x10) {
  return _retryQuery.apply(this, arguments);
}
function _retryQuery() {
  _retryQuery = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee7(elementId, type) {
    var inputElement, outputElement, gptResponse, equalIndex, evaluatedText;
    return _regeneratorRuntime().wrap(function _callee7$(_context7) {
      while (1) switch (_context7.prev = _context7.next) {
        case 0:
          if (!(elementId === undefined)) {
            _context7.next = 2;
            break;
          }
          return _context7.abrupt("return");
        case 2:
          if (!(type === undefined)) {
            _context7.next = 4;
            break;
          }
          return _context7.abrupt("return");
        case 4:
          inputElement = document.getElementById("input-" + elementId);
          outputElement = document.getElementById("output-" + elementId); // Clear output div
          outputElement.innerHTML = "";

          // Initiate Loader
          loader(outputElement);
          gptResponse = ""; // Get new GPT text
          if (!(type === "formula")) {
            _context7.next = 15;
            break;
          }
          _context7.next = 12;
          return getFormula(inputElement.textContent.trim());
        case 12:
          gptResponse = _context7.sent.trim();
          _context7.next = 22;
          break;
        case 15:
          if (!(type === "explain")) {
            _context7.next = 21;
            break;
          }
          _context7.next = 18;
          return getExplain(inputElement.textContent.trim());
        case 18:
          gptResponse = _context7.sent.trim();
          _context7.next = 22;
          break;
        case 21:
          return _context7.abrupt("return");
        case 22:
          // Check if '=' exists in the gptResponse
          equalIndex = gptResponse.indexOf("=");
          if (equalIndex !== -1) {
            // Remove text to the left of '='
            evaluatedText = gptResponse.slice(equalIndex).trim();
            clearInterval(loadInterval);
            outputElement.innerHTML = "";

            // Create HTML element from GPT text
            typeText(outputElement, wrapCellReferencesWithSpan(evaluatedText));
          } else {
            clearInterval(loadInterval);
            outputElement.innerHTML = "";

            // Create HTML element from GPT text
            typeText(outputElement, wrapCellReferencesWithSpan(gptResponse));
          }
        case 24:
        case "end":
          return _context7.stop();
      }
    }, _callee7);
  }));
  return _retryQuery.apply(this, arguments);
}
function copyToClipboard(text) {
  // Create a temporary textarea element
  var textarea = document.createElement("textarea");
  textarea.value = text;

  // Append the textarea to the document body
  document.body.appendChild(textarea);

  // Select the text within the textarea
  textarea.select();

  // Copy the selected text to the clipboard
  document.execCommand("copy");

  // Remove the temporary textarea
  document.body.removeChild(textarea);
}
function insertFormula(buttonId) {
  var messageId = buttonId.replace("insert-", "output-");
  var messageElement = document.getElementById(messageId);
  copyToClipboard(messageElement.textContent.trim()); //copy to user clipboard just in case

  Excel.run(function (context) {
    // Load the selected range and its properties
    var range = context.workbook.getSelectedRange();
    range.load("rowCount");
    range.load("columnCount");
    return context.sync().then(function () {
      // Check if only one cell is selected
      if (range.rowCount === 1 && range.columnCount === 1) {
        // Set the formula in the current cell
        range.formulas = [[messageElement.textContent.trim()]];
      }
    }).then(context.sync);
  }).catch(function (error) {
    console.log(error);
  });
}
function copyElementToClipboard(elementId) {
  var element = document.getElementById(elementId);
  copyToClipboard(element.textContent.trim());
}
//#endregion

//#region Handle prompt inputs and buttons
function autoResizeTextarea(textareaIds) {
  textareaIds.forEach(function (textareaId) {
    var textarea = document.getElementById(textareaId);
    var minHeight = getComputedStyle(textarea).height.replace(/px$/, "");
    var maxHeight = getComputedStyle(textarea).maxHeight.replace(/px$/, "");
    textarea.addEventListener("input", function (event) {
      // Calculate the scroll height of the textarea content
      textarea.style.height = "1px";
      var scrollHeight = textarea.scrollHeight;

      // Check if the scroll height exceeds the maximum height
      if (scrollHeight > maxHeight) {
        textarea.style.overflowY = "scroll"; // Display the scrollbar
        textarea.style.height = maxHeight + "px"; // Set the height to the maximum height
      } else if (scrollHeight < minHeight) {
        textarea.style.overflowY = "hidden"; // Hide the scrollbar
        textarea.style.height = minHeight + "px"; // Set the height to fit the content
      } else {
        textarea.style.overflowY = "hidden"; // Hide the scrollbar
        textarea.style.height = scrollHeight + "px"; // Set the height to fit the content
      }
    });
  });
}

// Format textarea elements
document.addEventListener("DOMContentLoaded", function () {
  autoResizeTextarea(["generateInput", "explainInput", "vbaInput"]); // Pass an array of textarea IDs
});

function submitOnEnter(elementIds) {
  elementIds.forEach(function (id) {
    var element = document.getElementById(id);
    if (element) {
      element.addEventListener("keydown", function (event) {
        if (event.key === "Enter" && !event.shiftKey) {
          event.preventDefault();
          if (id === "generateInput") {
            createFormula(id);
          } else if (id === "explainInput") {
            explainFormula(id);
          } else if (id === "vbaInput") {
            vbaFormula(id);
          }
        }
      });
    }
  });
}
var inputIds = ["generateInput", "explainInput", "vbaInput"]; // Provide the element IDs as an array
submitOnEnter(inputIds);
function getCurrent() {
  Excel.run(function (context) {
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load("formulas");
    return context.sync().then(function () {
      if (selectedRange.formulas.length === 1) {
        var formula = selectedRange.formulas[0][0];
        var formulaElement = document.getElementById("explainInput");
        formulaElement.textContent = formula;
      }
    });
  }).catch(function (error) {
    console.error(error);
  });
}
var generateHTML, explainHTML, vbaHTML;
document.addEventListener("DOMContentLoaded", function (event) {
  generateHTML = document.getElementById("output").innerHTML;
  explainHTML = document.getElementById("explain-output").innerHTML;
  vbaHTML = document.getElementById("code-output").innerHTML;
});
function refreshPage(id, htmlTag) {
  output = document.getElementById(id);
  output.innerHTML = htmlTag;
  if (htmlTag === explainHTML) {
    initializeExplainButtons();
  }
  if (htmlTag === generateHTML) {
    initializeGenerateButtons();
  }
}

//#endregion

//#region Functions to handle server querys || TODO: Move to separate file

// Process user input from textarea elements
function processInput(id) {
  // Get text from div
  var textInput = document.getElementById(id);
  if (textInput.tagName === "TEXTAREA") {
    var text = textInput.value.trim();
  } else if (textInput.tagName === "DIV") {
    var text = textInput.textContent.trim();
  } else {
    return;
  }
  if (text === "") {
    return;
  } else {
    // Clear the form after text entry
    textInput.value = "";
    return text;
  }
}

// Create a formula from user input
function getFormula(_x11) {
  return _getFormula.apply(this, arguments);
} // Explain a formula from user input
function _getFormula() {
  _getFormula = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee8(query) {
    var response, data;
    return _regeneratorRuntime().wrap(function _callee8$(_context8) {
      while (1) switch (_context8.prev = _context8.next) {
        case 0:
          _context8.prev = 0;
          _context8.next = 3;
          return fetch("https://testing-f03s.onrender.com/formula", {
            method: "POST",
            headers: {
              "Content-Type": "application/json"
            },
            body: JSON.stringify({
              prompt: query,
              token: auth
            })
          });
        case 3:
          response = _context8.sent;
          if (!response.ok) {
            _context8.next = 11;
            break;
          }
          _context8.next = 7;
          return response.json();
        case 7:
          data = _context8.sent;
          return _context8.abrupt("return", data.bot);
        case 11:
          console.error("Error:", response.status);
        case 12:
          _context8.next = 17;
          break;
        case 14:
          _context8.prev = 14;
          _context8.t0 = _context8["catch"](0);
          console.error("Error:", _context8.t0);
        case 17:
        case "end":
          return _context8.stop();
      }
    }, _callee8, null, [[0, 14]]);
  }));
  return _getFormula.apply(this, arguments);
}
function getExplain(_x12) {
  return _getExplain.apply(this, arguments);
} // Get VBA file from user input
function _getExplain() {
  _getExplain = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee9(query) {
    var response, data;
    return _regeneratorRuntime().wrap(function _callee9$(_context9) {
      while (1) switch (_context9.prev = _context9.next) {
        case 0:
          _context9.prev = 0;
          _context9.next = 3;
          return fetch("https://testing-f03s.onrender.com/explain", {
            method: "POST",
            headers: {
              "Content-Type": "application/json"
            },
            body: JSON.stringify({
              prompt: query,
              token: auth
            })
          });
        case 3:
          response = _context9.sent;
          if (!response.ok) {
            _context9.next = 11;
            break;
          }
          _context9.next = 7;
          return response.json();
        case 7:
          data = _context9.sent;
          return _context9.abrupt("return", data.bot);
        case 11:
          console.error("Error:", response.status);
        case 12:
          _context9.next = 17;
          break;
        case 14:
          _context9.prev = 14;
          _context9.t0 = _context9["catch"](0);
          console.error("Error:", _context9.t0);
        case 17:
        case "end":
          return _context9.stop();
      }
    }, _callee9, null, [[0, 14]]);
  }));
  return _getExplain.apply(this, arguments);
}
function getVBA(_x13) {
  return _getVBA.apply(this, arguments);
}
function _getVBA() {
  _getVBA = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee10(query) {
    var response, data;
    return _regeneratorRuntime().wrap(function _callee10$(_context10) {
      while (1) switch (_context10.prev = _context10.next) {
        case 0:
          _context10.prev = 0;
          _context10.next = 3;
          return fetch("https://testing-f03s.onrender.com/vba", {
            method: "POST",
            headers: {
              "Content-Type": "application/json"
            },
            body: JSON.stringify({
              prompt: query,
              token: auth
            })
          });
        case 3:
          response = _context10.sent;
          if (!response.ok) {
            _context10.next = 11;
            break;
          }
          _context10.next = 7;
          return response.json();
        case 7:
          data = _context10.sent;
          return _context10.abrupt("return", data.bot);
        case 11:
          console.error("Error:", response.status);
        case 12:
          _context10.next = 17;
          break;
        case 14:
          _context10.prev = 14;
          _context10.t0 = _context10["catch"](0);
          console.error("Error:", _context10.t0);
        case 17:
        case "end":
          return _context10.stop();
      }
    }, _callee10, null, [[0, 14]]);
  }));
  return _getVBA.apply(this, arguments);
}
function createFormula(_x14) {
  return _createFormula.apply(this, arguments);
}
function _createFormula() {
  _createFormula = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee11(id) {
    var text, divElement, uniqueId, wrappedText, outputElement, output, messageDiv, gptResponse, equalIndex, evaluatedText;
    return _regeneratorRuntime().wrap(function _callee11$(_context11) {
      while (1) switch (_context11.prev = _context11.next) {
        case 0:
          text = processInput(id); // Clear output div
          divElement = document.getElementById("output");
          divElement.innerHTML = "";
          uniqueId = generateUniqueId();
          wrappedText = wrapCellReferencesWithSpan(text); // Create HTML element from user text
          outputElement = document.getElementById("output");
          output = outputText(false, wrappedText, uniqueId);
          outputElement.innerHTML += output;

          // Pre-create HTML element for GPT text
          outputElement.innerHTML += outputText(true, "", uniqueId);

          // Get the pre-created HTML element and add loader element
          messageDiv = document.getElementById("output-" + uniqueId);
          loader(messageDiv);
          _context11.next = 13;
          return getFormula(text);
        case 13:
          gptResponse = _context11.sent.trim();
          // Get GPT text
          // Check if '=' exists in the gptResponse
          equalIndex = gptResponse.indexOf("=");
          if (equalIndex !== -1) {
            // Remove text to the left of '='
            evaluatedText = gptResponse.slice(equalIndex).trim();
            clearInterval(loadInterval);
            messageDiv.innerHTML = "";

            // Create HTML element from GPT text
            typeText(messageDiv, wrapCellReferencesWithSpan(evaluatedText), uniqueId, "formula");
          } else {
            clearInterval(loadInterval);
            messageDiv.innerHTML = "";

            // Create HTML element from GPT text
            typeText(messageDiv, wrapCellReferencesWithSpan(gptResponse), uniqueId, "formula");
          }
        case 16:
        case "end":
          return _context11.stop();
      }
    }, _callee11);
  }));
  return _createFormula.apply(this, arguments);
}
function explainFormula(_x15) {
  return _explainFormula.apply(this, arguments);
}
function _explainFormula() {
  _explainFormula = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee12(id) {
    var text, textInput, divElement, uniqueId, wrappedText, outputElement, output, messageDiv, gptResponse, equalIndex, evaluatedText;
    return _regeneratorRuntime().wrap(function _callee12$(_context12) {
      while (1) switch (_context12.prev = _context12.next) {
        case 0:
          text = processInput(id); // Get text from div
          textInput = document.getElementById(id); // Clear output div
          divElement = document.getElementById("explain-output");
          divElement.innerHTML = "";
          uniqueId = generateUniqueId();
          wrappedText = wrapCellReferencesWithSpan(text); // Create HTML element from user text
          outputElement = document.getElementById("explain-output");
          output = outputText(false, wrappedText, uniqueId);
          outputElement.innerHTML += output;

          // Clear the form after text entry
          textInput.textContent = "";

          // Pre-create HTML element for GPT text
          outputElement.innerHTML += outputText(true, "", uniqueId);

          // Get the pre-created HTML element and add loader element
          messageDiv = document.getElementById("output-" + uniqueId);
          loader(messageDiv);
          _context12.next = 15;
          return getExplain(text);
        case 15:
          gptResponse = _context12.sent.trim();
          // Get GPT text
          // Check if '=' exists in the gptResponse
          equalIndex = gptResponse.indexOf("=");
          if (equalIndex !== -1) {
            // Remove text to the left of '='
            evaluatedText = gptResponse.slice(equalIndex).trim();
            clearInterval(loadInterval);
            messageDiv.innerHTML = "";

            // Create HTML element from GPT text
            typeText(messageDiv, wrapCellReferencesWithSpan(evaluatedText), uniqueId, "explain");
          } else {
            clearInterval(loadInterval);
            messageDiv.innerHTML = "";

            // Create HTML element from GPT text
            typeText(messageDiv, wrapCellReferencesWithSpan(gptResponse), uniqueId, "explain");
          }
        case 18:
        case "end":
          return _context12.stop();
      }
    }, _callee12);
  }));
  return _explainFormula.apply(this, arguments);
}
function vbaFormula(_x16) {
  return _vbaFormula.apply(this, arguments);
} //#endregion
//#region Functions to handle text and code formatting
function _vbaFormula() {
  _vbaFormula = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee13(id) {
    var text, outputElement, gptResponse, regex, match, result;
    return _regeneratorRuntime().wrap(function _callee13$(_context13) {
      while (1) switch (_context13.prev = _context13.next) {
        case 0:
          text = processInput(id); // Output div
          outputElement = document.getElementById("code-output"); // Clear output div and text input
          outputElement.innerHTML = "";
          loader(outputElement);
          _context13.next = 6;
          return getVBA(text);
        case 6:
          gptResponse = _context13.sent;
          regex = /```([\s\S]*?)```/;
          match = gptResponse.match(regex);
          if (match) {
            result = match[1]; // Return the captured group
          } else {
            result = gptResponse;
          }
          clearInterval(loadInterval);
          outputElement.innerHTML = "";
          result = formatCodeWithHLJS(result);
          typeText(outputElement, result);
        case 14:
        case "end":
          return _context13.stop();
      }
    }, _callee13);
  }));
  return _vbaFormula.apply(this, arguments);
}
function wrapCellReferencesWithSpan(text) {
  var cellReferenceRegex = /\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?/g;
  var cellReferences = text.match(cellReferenceRegex);
  var colors = ["#00A7E1", "#4094DE", "#6A7ED1", "#8A65BA", "#9F4999", "#A82B70"]; // List of colors
  var colorMap = {}; // Map to track assigned colors

  var wrappedText = text.replace(cellReferenceRegex, function (match) {
    if (!colorMap[match]) {
      var colorIndex = Object.keys(colorMap).length % colors.length;
      colorMap[match] = colors[colorIndex];
    }
    var spanStyle = "color: ".concat(colorMap[match]);
    return "<span class=\"cell-reference\" style=\"".concat(spanStyle, "\">").concat(match, "</span>");
  });
  return wrappedText;
}
var hoveredElement = null;
function handleHover(event) {
  hoveredElement = event.target;
}

// Function to toggle the cell reference format
function toggleCellReferenceFormat() {
  if (hoveredElement) {
    var currentText = hoveredElement.textContent;
    var newText = "";

    // Regular expressions for pattern matching
    var rangePattern = /^([$]?[A-Z]+)([$]?[0-9]+):([$]?[A-Z]+)([$]?[0-9]+)$/;
    var cellPattern = /^([$]?[A-Z]+)([$]?[0-9]+)$/;

    // Check if it's a range
    if (rangePattern.test(currentText)) {
      newText = currentText.replace(rangePattern, function (match, startCol, startRow, endCol, endRow) {
        if (startRow.startsWith("$")) {
          if (!endCol.startsWith("$")) {
            // Remove '$' at the start if present
            startRow = startRow.substr(1);
            endRow = endRow.substr(1);
          }
        } else {
          if (!endCol.startsWith("$")) {
            // Add '$' at the start
            startRow = "$" + startRow;
            endRow = "$" + endRow;
          }
        }
        if (startCol.startsWith("$")) {
          // Remove '$' at the start if present
          startCol = startCol.substr(1);
          endCol = endCol.substr(1);
        } else {
          // Add '$' at the start
          startCol = "$" + startCol;
          endCol = "$" + endCol;
        }
        return startCol + startRow + ":" + endCol + endRow;
      });
    }
    // Check if it's a single cell
    else if (cellPattern.test(currentText)) {
      newText = currentText.replace(cellPattern, function (match, col, row) {
        if (row.startsWith("$")) {
          if (!col.startsWith("$")) {
            // Remove '$' at the start if present
            row = row.substr(1);
          }
        } else {
          if (!col.startsWith("$")) {
            // Add '$' at the start
            row = "$" + row;
          }
        }
        if (col.startsWith("$")) {
          // Remove '$' at the start if present
          col = col.substr(1);
        } else {
          // Add '$' at the start
          col = "$" + col;
        }
        return col + row;
      });
    }

    // Update the text of the hovered element
    hoveredElement.textContent = newText;
  }
}

// Function to handle the F4 key press event
function handleKeyPress(event) {
  // Check if the pressed key is F4 (keyCode 115) and hovered element is a cell reference
  if (event.keyCode === 115 && hoveredElement && hoveredElement.classList.contains("cell-reference")) {
    toggleCellReferenceFormat();
  }
}
// Register the event listeners
document.addEventListener("mouseover", handleHover);
document.addEventListener("keydown", handleKeyPress);
function formatCodeWithHLJS(code) {
  var tempElement = document.createElement("div");
  tempElement.innerHTML = code;
  hljs.highlightBlock(tempElement);
  return tempElement.innerHTML;
}

// Handle code formatting on the page
function formatCode() {
  document.addEventListener("DOMContentLoaded", function (event) {
    hljs.highlightAll();
  });
}
formatCode();
//#endregion

//#region Button Clicks || TODO: Debug each button

// –––––––––– Example buttons –––––––––––

function initializeGenerateButtons() {
  document.getElementById("button-rec1").addEventListener("click", function () {
    // Create: Example
    createFormula("rec1");
  });
  document.getElementById("button-rec2").addEventListener("click", function () {
    // Create: Example
    createFormula("rec2");
  });
  document.getElementById("button-rec3").addEventListener("click", function () {
    // Create: Example
    createFormula("rec3");
  });
}
initializeGenerateButtons();
function initializeExplainButtons() {
  document.getElementById("explain1").addEventListener("click", function () {
    // Explain: Example
    explainFormula("explain1");
  });
  document.getElementById("explain2").addEventListener("click", function () {
    // Explain: Example
    explainFormula("explain2");
  });
  document.getElementById("explain3").addEventListener("click", function () {
    // Explain: Example
    explainFormula("explain3");
  });
  document.getElementById("explain4").addEventListener("click", function () {
    // Explain: Example
    explainFormula("explain4");
  });
  document.getElementById("explain5").addEventListener("click", function () {
    // Explain: Example
    explainFormula("explain5");
  });
  document.getElementById("explain6").addEventListener("click", function () {
    // Explain: Example
    explainFormula("explain6");
  });
}
initializeExplainButtons();

// –––––––––– Input buttons –––––––––––
document.getElementById("generateInput-button").addEventListener("click", function () {
  // Create formula button
  createFormula("generateInput");
});
document.getElementById("button-explainInput").addEventListener("click", function () {
  // Explain formula button
  explainFormula("explainInput");
});
document.getElementById("button-vbaInput").addEventListener("click", function () {
  // VBA formula button
  vbaFormula("vbaInput");
});
document.getElementById("generate-refresh").addEventListener("click", function () {
  // Generate: Refresh App
  refreshPage("output", generateHTML);
});
document.getElementById("explain-refresh").addEventListener("click", function () {
  // Explain: Refresh App
  refreshPage("explain-output", explainHTML);
});
document.getElementById("vba-refresh").addEventListener("click", function () {
  // VBA: Refresh App
  refreshPage("code-output", vbaHTML);
});
document.getElementById("explain-current").addEventListener("click", function () {
  // Explain: Get Current
  getCurrent();
});
document.getElementById("vba-current").addEventListener("click", function () {
  // VBA: Get Current
  getCurrent();
});

// –––––––––– Output buttons –––––––––––
document.getElementById("vba-copy").addEventListener("click", function () {
  // VBA: Copy to Clipboard
  copyElementToClipboard("code-output");
});

// –––––––––– Authentication buttons –––––––––––
var emailTextarea = document.getElementById("signup-email");
emailTextarea.addEventListener("focus", function () {
  var placeholder = emailTextarea.placeholder;
  if (emailTextarea.value === "") {
    emailTextarea.value = placeholder;
  }
});
emailTextarea.addEventListener("blur", function () {
  var placeholder = emailTextarea.placeholder;
  if (emailTextarea.value === placeholder) {
    emailTextarea.value = "";
  }
});
document.getElementById("auth-button").addEventListener("click", function () {
  authenticateUser("signup-email", "auth-status");
});
document.getElementById("deauth-button").addEventListener("click", function () {
  deAuthenticateUser("signup-email", "auth-status");
});
//#endregion