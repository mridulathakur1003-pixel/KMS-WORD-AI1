# ================================================================
# KMS WORD AI — VERSION 20.0 PRO
# Advanced Office-style document creation suite
# Python 3 + PyQt6 + Qt WebEngine
#
# Features:
# Dashboard, document manager, rich-text editor, ribbon tabs,
# formatting, tables, images, links, page layout, themes,
# navigation/outline, find & replace, comments, word statistics,
# autosave, local recovery, HTML/TXT export, print/PDF, whiteboard,
# templates, focus mode, zoom, keyboard shortcuts and dark workspace.
#
# Note: This is a desktop document editor, not a Microsoft Word clone.
# DOCX import/export and cloud collaboration are not included.
# ================================================================

import sys
import os
import json
import uuid
import hashlib
import urllib.request
import urllib.error
from pathlib import Path
from PyQt6.QtCore import QUrl, QObject, pyqtSlot
from PyQt6.QtGui import QAction, QKeySequence
from PyQt6.QtWidgets import (QApplication, QMainWindow, QMessageBox, QFileDialog,
                             QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                             QPushButton, QFrame)
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
from PyQt6.QtWebChannel import QWebChannel
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebEngineCore import QWebEngineSettings


HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>KMS WORD AI V20.0 PRO</title>
<style>
:root{
 --navy:#14243d;--navy2:#1d3558;--blue:#245bb5;--blue2:#3678dc;
 --saffron:#ed8b2f;--green:#16834c;--bg:#e9edf4;--panel:#f8fafd;
 --paper:#fff;--ink:#202938;--muted:#68778d;--line:#d8e0eb;
 --shadow:0 8px 28px rgba(20,35,59,.12);--accent:#245bb5
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:14px "Segoe UI",Arial,sans-serif;overflow:hidden}
button,input,select,textarea{font:inherit}
button{border:1px solid var(--line);background:#fff;color:#273750;border-radius:6px;padding:7px 10px;cursor:pointer}
button:hover{background:#edf3fc;border-color:#b9cbea}
button:disabled{opacity:.45;cursor:not-allowed}
button.primary{background:var(--blue);border-color:var(--blue);color:#fff}
button.primary:hover{background:var(--blue2)}
button.warn{color:#a33}
input,select,textarea{border:1px solid var(--line);border-radius:6px;padding:8px;background:#fff;color:#26364d}
.topbar{height:55px;background:var(--navy);color:white;display:flex;align-items:center;gap:12px;padding:0 16px}
.brand{font-weight:800;font-size:19px;white-space:nowrap;letter-spacing:.2px}
.brand em{font-style:normal;color:#ffb15c}
.version{font-size:10px;color:#b8c8dd;margin-left:3px}
.doc-title{width:220px;background:#263c5d;border:1px solid #4c6382;color:white}
.top-spacer{flex:1}
.top-search{width:min(290px,25vw)}
.topbar button{background:#2a4163;color:white;border-color:#49617f}
.topbar button:hover{background:#39577e}
.avatar{width:31px;height:31px;border-radius:50%;display:grid;place-items:center;background:#f0a54d;color:#15243c;font-weight:800}
.tabs{height:40px;background:white;border-bottom:1px solid var(--line);display:flex;padding-left:10px;gap:3px;overflow-x:auto}
.tab{border:0;border-radius:5px 5px 0 0;background:transparent;padding:0 15px;font-weight:650;white-space:nowrap}
.tab.active{color:var(--accent);border-bottom:3px solid var(--accent);background:#f3f7fd}
.ribbon{height:111px;background:#f9fbfe;border-bottom:1px solid var(--line);display:flex;align-items:stretch;padding:8px 10px;gap:4px;overflow-x:auto}
.rgroup{display:flex;align-items:center;gap:5px;padding:2px 9px;border-right:1px solid #dce3ed;white-space:nowrap}
.rgroup:last-child{border:0}
.rgroup-inner{display:flex;gap:4px;align-items:center;flex-wrap:wrap;max-width:370px}
.rlabel{font-size:10px;color:#7b8798;text-align:center;margin-top:5px}
.rtool{display:flex;flex-direction:column;align-items:center;gap:3px;min-width:40px}
.rtool button{height:32px;min-width:35px;padding:4px 7px;font-size:14px}
.rtool small{font-size:10px;color:#58677d}
.ribbon select,.ribbon input{height:30px;padding:3px 6px;max-width:160px}
.font-name{width:135px}.font-size{width:58px}
.workspace{height:calc(100vh - 238px);display:flex;min-height:250px}
.sidebar{width:205px;background:var(--panel);border-right:1px solid var(--line);padding:12px 10px;overflow:auto}
.side-heading{font-size:10px;font-weight:800;letter-spacing:1.2px;color:#8995a5;margin:12px 9px 6px}
.side-item{display:flex;align-items:center;gap:10px;padding:9px 10px;border-radius:7px;cursor:pointer;color:#43536b;font-size:13px}
.side-item:hover,.side-item.active{background:#e7effc;color:#174f9e}
.side-item span:first-child{width:21px;text-align:center;font-size:16px}
.mainarea{flex:1;display:flex;min-width:0;overflow:hidden}
.editor-scroll{flex:1;overflow:auto;padding:27px 25px 50px;display:flex;justify-content:center;align-items:flex-start}
.page-shell{width:816px;min-width:650px;position:relative}
.page{width:816px;min-height:1056px;background:var(--paper);color:#222;padding:82px 80px;margin:0 auto 28px;box-shadow:var(--shadow);outline:none;line-height:1.55;font-family:Arial,sans-serif;font-size:11pt;overflow-wrap:anywhere}
.page:focus{outline:1px solid #c8d8f0}
.page-border-double{border:5px double #294d7a}.page-border-shadow{border:1px solid #9aa9bd;box-shadow:0 0 0 7px #e7edf6,0 12px 32px rgba(20,35,59,.22)}.page-border-dashed{border:3px dashed #567da8}.page-border-rounded{border:4px solid #2b659f;border-radius:18px}.page-border-classic{border:2px solid #333;outline:1px solid #999;outline-offset:-9px}.page-border-gold{border:7px double #b58a36}.page-border-none{border:0}
.temp-entry-highlight{background:#fff09a;transition:background-color 5s ease}
.page-header{border-bottom:1px solid #9aa9bd;padding:0 0 10px;margin:0 0 24px;color:#44536a;font-size:10pt;min-height:24px}
.page-footer{border-top:1px solid #9aa9bd;padding:10px 0 0;margin:26px 0 0;color:#44536a;font-size:9pt;min-height:22px}
.page-header:empty:before{content:'Header — click to edit';color:#9aa4b2}.page-footer:empty:before{content:'Footer — click to edit';color:#9aa4b2}
.page h1,.page h2,.page h3{line-height:1.25}
.page table{border-collapse:collapse;width:100%;margin:12px 0}
.page td,.page th{border:1px solid #777;padding:7px;min-width:32px}
.page img{max-width:100%;height:auto}
.page blockquote{border-left:4px solid var(--accent);padding:8px 16px;margin:12px 0;background:#f4f7fc;color:#46546a}
.page hr{border:0;border-top:1px solid #aaa;margin:18px 0}

.page .callout{padding:14px 16px;border-left:5px solid var(--accent);background:#f2f6fc;margin:14px 0;border-radius:4px}
.page .cover-kicker{letter-spacing:3px;text-transform:uppercase;font-size:10pt;color:#6b7890}
.page .cover-title{font-size:34pt;line-height:1.12;margin:18px 0 10px}
.page .cover-subtitle{font-size:15pt;color:#64748b}
.page .cover-meta{margin-top:90px;padding-top:14px;border-top:1px solid #ccd5e2;color:#58677d}
.page .section-band{padding:9px 13px;background:#eaf0f8;border-left:4px solid var(--accent);font-weight:700}
.page .note-box{border:1px solid #d8e0eb;background:#fbfcff;padding:12px;border-radius:7px}
.page .table-blue th{background:#e7effb;color:#173e70}
.page .table-green th{background:#e6f3ec;color:#185c3b}
.page .table-dark th{background:#29384f;color:#fff}
.page .table-minimal td,.page .table-minimal th{border-left:0;border-right:0;border-top:0}
.page .watermark{position:absolute;pointer-events:none;opacity:.10;transform:rotate(-32deg);font-size:48pt;font-weight:800;top:46%;left:20%;z-index:0}
.page .page-number{display:inline-block;min-width:24px;text-align:center}
.template-card{background:#fff;border:1px solid var(--line);border-radius:10px;padding:14px;min-height:145px}
.template-card h3{margin:0 0 7px;color:var(--navy)}
.template-card p{font-size:12px;color:var(--muted);min-height:34px}
.theme-swatch{height:10px;border-radius:8px;margin:9px 0}

.gallery-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(125px,1fr));gap:8px;margin:12px 0}
.gallery-item{min-height:68px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:5px;border:1px solid #d8e0eb;border-radius:8px;background:#fff;padding:8px;cursor:pointer;color:#25364e}
.gallery-item:hover{border-color:#3678dc;background:#edf4ff}
.gallery-preview{font-size:25px;font-weight:800;line-height:1.2}
.gallery-item small{font-size:10px;text-align:center;color:#58677d}


.rightpanel{width:230px;background:var(--panel);border-left:1px solid var(--line);padding:12px;overflow:auto}
.side-hide{width:100%;margin:0 0 10px;padding:6px 8px;font-size:11px;background:#edf3fc}
body.hide-left #leftSidebar,body.hide-right #rightSidebar{display:none!important}
body.hide-left .editor-scroll,body.hide-right .editor-scroll{padding-left:14px;padding-right:14px}
.sidebar-reopen{display:none;position:fixed;z-index:45;top:205px;background:#17365f;color:#fff;border:0;border-radius:0 9px 9px 0;padding:12px 9px;box-shadow:0 4px 16px #10213b35;writing-mode:vertical-rl;letter-spacing:.3px}
#reopenLeft{left:0}#reopenRight{right:0;border-radius:9px 0 0 9px}
body.hide-left #reopenLeft,body.hide-right #reopenRight{display:block}
.page .kms-shape{display:inline-flex;position:relative;align-items:center;justify-content:center;vertical-align:middle;min-width:90px;min-height:65px;width:150px;height:88px;margin:9px;padding:12px;background:#dbeafe;color:#174b91;border:3px solid #245bb5;font-size:18px;font-weight:700;text-align:center;white-space:normal;overflow:visible;resize:both;cursor:move;user-select:text;outline:1px dashed transparent;filter:drop-shadow(1px 0 0 var(--shape-border,#245bb5)) drop-shadow(-1px 0 0 var(--shape-border,#245bb5)) drop-shadow(0 1px 0 var(--shape-border,#245bb5)) drop-shadow(0 -1px 0 var(--shape-border,#245bb5))}
.page .kms-shape:focus,.page .kms-shape.object-selected{outline:2px solid #3678dc;outline-offset:3px}
.page .kms-shape[data-shape="Circle"],.page .kms-shape[data-shape="Ring"]{border-radius:50%;aspect-ratio:1/1}
.page .kms-shape[data-shape="Oval"]{border-radius:50%}
.page .kms-shape[data-shape="Triangle"]{clip-path:polygon(50% 0,0 100%,100% 100%);border:0;border-radius:0}
.page .kms-shape[data-shape="Inverted triangle"]{clip-path:polygon(0 0,100% 0,50% 100%);border:0}
.page .kms-shape[data-shape="Diamond"],.page .kms-shape[data-shape="Flowchart decision"]{clip-path:polygon(50% 0,100% 50%,50% 100%,0 50%);border:0}
.page .kms-shape[data-shape="Hexagon"],.page .kms-shape[data-shape="Hexagon outline"]{clip-path:polygon(25% 0,75% 0,100% 50%,75% 100%,25% 100%,0 50%);border:0}
.page .kms-shape[data-shape="Pentagon"]{clip-path:polygon(50% 0,100% 38%,82% 100%,18% 100%,0 38%);border:0}
.page .kms-shape[data-shape="Parallelogram"]{clip-path:polygon(20% 0,100% 0,80% 100%,0 100%);border:0}
.page .kms-shape[data-shape="Trapezoid"]{clip-path:polygon(20% 0,80% 0,100% 100%,0 100%);border:0}
.page .kms-shape[data-shape="Octagon"]{clip-path:polygon(28% 0,72% 0,100% 28%,100% 72%,72% 100%,28% 100%,0 72%,0 28%);border:0}
.page img.kms-image{display:inline-block;max-width:100%;height:auto;resize:both;overflow:hidden;outline:1px dashed transparent;cursor:move;vertical-align:middle;object-fit:contain}
.page img.kms-image.object-selected{outline:2px solid #3678dc;outline-offset:3px}
#whiteboard{display:none;position:fixed;inset:0;z-index:70;background:#eaf0f7;flex-direction:column}
.boardbar{min-height:54px;background:#14243d;display:flex;align-items:center;gap:6px;padding:7px 10px;color:white;flex-wrap:wrap;overflow:auto}
.boardbar button{font-size:12px;padding:6px 8px;white-space:nowrap}
.boardbar input[type=color]{width:34px;height:30px;padding:2px}
.boardstage{flex:1;min-height:0;position:relative;overflow:hidden;background:#eaf0f7}
#board{display:block;width:100%;height:100%;touch-action:none;cursor:crosshair;background:#fff}
.boardstatus{height:26px;background:#14243d;color:#dce8f8;padding:4px 12px;font-size:11px;display:flex;gap:18px}

.page .kms-shape{display:inline-flex;align-items:center;justify-content:center;min-width:76px;min-height:54px;padding:8px 13px;margin:5px;border:2px solid #245bb5;border-radius:8px;background:#eaf1fc;color:#174b91;font-size:25px;font-weight:700;vertical-align:middle;resize:both;overflow:auto;outline:1px dashed transparent;cursor:text}
.page .kms-shape{position:relative;display:inline-flex;align-items:center;justify-content:center;min-width:90px;min-height:65px;padding:10px 16px;margin:7px;border:2px solid #245bb5;background:#eaf1fc;color:#174b91;font-size:22px;font-weight:700;vertical-align:middle;resize:both;overflow:auto;outline:1px dashed transparent;cursor:text;user-select:text;white-space:normal;text-align:center}
.page .kms-shape:focus,.page .kms-shape.object-selected{outline:2px solid #3678dc;outline-offset:2px}
.page .kms-shape[data-shape="Circle"]{border-radius:50%;aspect-ratio:1/1}
.page .kms-shape[data-shape="Oval"]{border-radius:50%}
.page .kms-shape[data-shape="Rounded rectangle"]{border-radius:18px}
.page .kms-shape[data-shape="Triangle"]{clip-path:polygon(50% 0,0 100%,100% 100%);border:0;padding-top:25px}
.page .kms-shape[data-shape="Diamond"]{clip-path:polygon(50% 0,100% 50%,50% 100%,0 50%);border:0}
.page .kms-shape[data-shape="Hexagon"]{clip-path:polygon(25% 0,75% 0,100% 50%,75% 100%,25% 100%,0 50%);border:0}
.page .kms-shape[data-shape="Pentagon"]{clip-path:polygon(50% 0,100% 38%,82% 100%,18% 100%,0 38%);border:0}
.page .kms-shape[data-shape="Parallelogram"]{clip-path:polygon(20% 0,100% 0,80% 100%,0 100%);border:0}
.page .kms-shape[data-shape="Trapezoid"]{clip-path:polygon(20% 0,80% 0,100% 100%,0 100%);border:0}
.page img.kms-image{display:inline-block;max-width:100%;height:auto;resize:both;overflow:auto;outline:1px dashed transparent;cursor:pointer;vertical-align:middle;object-fit:contain}
.page img.kms-image.object-selected{outline:2px solid #3678dc;outline-offset:2px}
#objectInspector input{width:100%;margin:3px 0 7px}#objectInspector button{margin:2px 2px 2px 0}
@page{margin:15mm}
@media print{
 @page{size:auto;margin:15mm}
 html,body{height:auto!important;overflow:visible!important;background:#fff!important}
 body{print-color-adjust:exact;-webkit-print-color-adjust:exact}
 .topbar,.tabs,.ribbon,.sidebar,.rightpanel,.status,.dashboard,.overlay,.toast,#whiteboard{display:none!important}
 .workspace,.mainarea,.editor-scroll,.page-shell{height:auto!important;display:block!important;overflow:visible!important;padding:0!important;margin:0!important;width:auto!important;min-width:0!important}
 .page{box-sizing:border-box;box-shadow:none!important;margin:0!important;width:100%!important;min-height:0!important;zoom:1!important;break-after:page;page-break-after:always;overflow:visible!important}
 .page:last-child{break-after:auto;page-break-after:auto}
 .kms-shape{resize:none!important}.object-selected{outline:none!important}
}

.pcard{background:white;border:1px solid var(--line);border-radius:9px;padding:12px;margin-bottom:11px}
.pcard h3{font-size:13px;margin:0 0 9px}
.pcard p{font-size:12px;color:var(--muted);line-height:1.45;margin:5px 0}
.pcard button{width:100%;text-align:left;margin:3px 0}
.status{height:32px;background:#f9fbfe;border-top:1px solid var(--line);display:flex;align-items:center;gap:16px;padding:0 12px;color:#66758a;font-size:11px}
.status .push{flex:1}
.status button{padding:3px 8px}
.overlay{display:none;position:fixed;inset:0;background:#101e34b8;z-index:50;align-items:center;justify-content:center;padding:18px}
.dialog{background:white;border-radius:12px;box-shadow:var(--shadow);width:min(650px,96vw);max-height:88vh;overflow:auto;padding:22px}
.dialog h2{margin:0 0 15px;color:var(--navy)}
.dialog label{font-size:12px;color:#596980;font-weight:650}
.dialog input,.dialog textarea,.dialog select{display:block;width:100%;margin:5px 0 12px}
.dialog-actions{display:flex;justify-content:flex-end;gap:8px;margin-top:14px}
.dashboard{display:none;position:fixed;inset:95px 0 32px;background:#f1f5fb;z-index:15;overflow:auto;padding:28px 34px}
.dash-head{display:flex;align-items:center;justify-content:space-between;gap:15px;flex-wrap:wrap}
.dash-head h1{margin:0;color:var(--navy);font-size:27px}
.dash-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(205px,1fr));gap:15px;margin-top:22px}
.dash-card{background:#fff;border:1px solid var(--line);border-radius:12px;padding:18px;box-shadow:0 4px 15px #1c31500c}
.dash-card .emoji{font-size:25px}
.dash-card h3{margin:8px 0 5px}.dash-card p{font-size:12px;color:var(--muted);line-height:1.5;min-height:36px}
.dash-card button{margin-top:8px}
.recent{background:#fff;border:1px solid var(--line);border-radius:12px;padding:17px;margin-top:18px}
.recent-row{display:flex;align-items:center;gap:10px;padding:9px 0;border-bottom:1px solid #edf0f5}
.recent-row:last-child{border:0}
.recent-row .grow{flex:1}
#whiteboard{display:none;position:fixed;inset:0;z-index:70;background:white;flex-direction:column}
.boardbar{height:52px;background:var(--navy);display:flex;align-items:center;gap:9px;padding:8px 14px;color:white}
#board{flex:1;touch-action:none;cursor:crosshair}
.toast{display:none;position:fixed;z-index:100;bottom:48px;left:50%;transform:translateX(-50%);background:#172943;color:white;padding:10px 17px;border-radius:8px;box-shadow:var(--shadow)}
#outlineList button{display:block;width:100%;text-align:left;margin:3px 0;border:0;background:transparent}
body.dark{--bg:#161e2c;--panel:#202b3d;--paper:#fff;--ink:#e6edf7;--line:#35445b}
body.dark .tabs,body.dark .ribbon,body.dark .status{background:#202b3d;color:#e6edf7}
body.dark .tab{color:#d5deeb}body.dark .tab.active{color:#8fbaff;background:#263650}
body.dark .sidebar,body.dark .rightpanel{background:#202b3d}
body.dark .side-item{color:#d0d9e8}body.dark .pcard{background:#29364a;color:#e6edf7}
body.dark .pcard p{color:#b4c1d3}body.dark .editor-scroll{background:#161e2c}
body.dark .dashboard{background:#182334;color:#e6edf7}body.dark .dash-card,body.dark .recent{background:#263449;color:#e6edf7}
@media(max-width:1100px){.rightpanel{display:none}.sidebar{width:175px}}
@media(max-width:760px){.sidebar{display:none}.top-search{display:none}.doc-title{width:130px}.page-shell,.page{width:650px}.page{padding:55px}.workspace{height:calc(100vh - 227px)}}
@media print{
 body{overflow:visible;background:white}.topbar,.tabs,.ribbon,.sidebar,.rightpanel,.status,.dashboard{display:none!important}
 .workspace,.mainarea,.editor-scroll{height:auto;display:block;overflow:visible;padding:0}
 .page{box-shadow:none;margin:0;width:100%;min-height:0;padding:18mm 20mm;page-break-after:always}
}
</style>
</head>
<body>
<header class="topbar">
 <div class="brand">▣ KMS <em>WORD AI</em><span class="version">VERSION 20.0 PRO</span></div>
 <input id="docTitle" class="doc-title" value="Untitled Document" title="Document title">
 <div class="top-spacer"></div>
 <input id="commandSearch" class="top-search" placeholder="Search commands… (try 'table')" onkeydown="if(event.key==='Enter')runCommandSearch()">
 <button onclick="newDocument()">＋ New</button><button onclick="saveDocument()">Save</button><button onclick="requestPrint()">🖨 Print</button>
 <button onclick="exportHTML()">Export</button><button onclick="openWritingAssistant()">✦ AI Assistant</button><div class="avatar">K</div>
</header>
<nav class="tabs" id="tabs">
 <button class="tab active" data-tab="home" onclick="setTab('home')">Home</button>
 <button class="tab" data-tab="insert" onclick="setTab('insert')">Insert</button>
 <button class="tab" data-tab="layout" onclick="setTab('layout')">Layout</button>
 <button class="tab" data-tab="design" onclick="setTab('design')">Design</button>
 <button class="tab" data-tab="references" onclick="setTab('references')">References</button>
 <button class="tab" data-tab="review" onclick="setTab('review')">Review</button>
 <button class="tab" data-tab="view" onclick="setTab('view')">View</button>
 <button class="tab" data-tab="help" onclick="setTab('help')">Help</button>
 <button class="tab" data-tab="filetools" onclick="setTab('filetools')">File</button>
 <button class="tab" data-tab="draw" onclick="setTab('draw')">Draw</button>
 <button class="tab" data-tab="ailab" onclick="setTab('ailab')">AI Lab</button>
 <button class="tab" data-tab="mailings" onclick="setTab('mailings')">Mailings</button>
 <button class="tab" data-tab="forms" onclick="setTab('forms')">Forms</button>
 <button class="tab" data-tab="data" onclick="setTab('data')">Data</button>
 <button class="tab" data-tab="security" onclick="setTab('security')">Security</button>
 <button class="tab" data-tab="accessibility" onclick="setTab('accessibility')">Accessibility</button>
 <button class="tab" data-tab="developer" onclick="setTab('developer')">Developer</button>
 <button class="tab" data-tab="automate" onclick="setTab('automate')">Automate</button>
</nav>
<section class="ribbon" id="ribbon"></section>
<div class="workspace" id="workspace">
 <aside class="sidebar" id="leftSidebar"><button class="side-hide" onclick="toggleSidebar('left')">◀ Hide sidebar</button>
  <div class="side-heading">WORKSPACE</div>
  <div class="side-item active" onclick="showDashboard(false)"><span>▤</span>Document editor</div>
  <div class="side-item" onclick="showDashboard(true)"><span>⌂</span>Dashboard</div>
  <div class="side-item" onclick="openWhiteboard()"><span>✎</span>Whiteboard</div>
  <div class="side-item" onclick="openFind()"><span>⌕</span>Find & replace</div>
  <div class="side-heading">DOCUMENT</div>
  <div class="side-item" onclick="insertTable()"><span>▦</span>Insert table</div>
  <div class="side-item" onclick="chooseImage()"><span>▧</span>Insert picture</div>
  <div class="side-item" onclick="requestPrint()"><span>▣</span>Print / PDF</div>
  <div class="side-item" onclick="showStatistics()"><span>≡</span>Statistics</div>
  <div class="side-heading">QUICK ACCESS</div>
  <div class="side-item" onclick="doFormat('undo')"><span>↶</span>Undo</div>
  <div class="side-item" onclick="doFormat('redo')"><span>↷</span>Redo</div>
  <div class="side-item" onclick="toggleFocus()"><span>⛶</span>Focus mode</div>
  <div class="side-item" onclick="toggleDark()"><span>◐</span>Dark workspace</div>
 </aside>
 <main class="mainarea">
  <section class="editor-scroll" id="editorScroll">
   <div class="page-shell"><article class="page" id="page" contenteditable="true" spellcheck="true">
    <h1 style="text-align:center;color:#174b91">KMS WORD AI — VERSION 20.0 PRO</h1>
    <p style="text-align:center;color:#e27e20"><b>Advanced document creation and productivity workspace</b></p>
    <hr>
    <p>Welcome to KMS WORD AI Version 14. This is your writing space. Click anywhere on this page and start creating.</p>
    <h2>Explore the workspace</h2>
    <ul>
     <li><b>Home:</b> font formatting, paragraph tools, styles and editing.</li>
     <li><b>Insert:</b> tables, images, links, shapes, symbols, date and page breaks.</li>
     <li><b>Layout:</b> page size, margins, orientation and spacing.</li>
     <li><b>Design:</b> document themes, page colour and borders.</li>
     <li><b>References:</b> contents, footnotes and citations.</li>
     <li><b>Review:</b> spelling support, comments, word count and replace.</li>
     <li><b>View:</b> focus mode, zoom, outline and whiteboard.</li>
    </ul>
    <p><b>KMS A VISION</b></p>
   </article></div>
  </section>
 </main>
 <aside class="rightpanel" id="rightSidebar"><button class="side-hide" onclick="toggleSidebar('right')">Hide panel ▶</button>
  <div class="pcard"><h3>Document insights</h3><p id="stats">Words: 0 · Characters: 0</p><button onclick="showStatistics()">Detailed statistics</button></div>
  <div class="pcard"><h3>Navigation / outline</h3><div id="outlineList"><p>Headings appear here.</p></div></div>
  <div class="pcard"><h3>Quick styles</h3><button onclick="applyBlock('h1')">Title / Heading 1</button><button onclick="applyBlock('h2')">Heading 2</button><button onclick="applyBlock('h3')">Heading 3</button><button onclick="applyBlock('p')">Normal text</button><button onclick="applyBlock('blockquote')">Quote</button></div>
  <div class="pcard"><h3>Page appearance</h3><button onclick="setPageColor('#ffffff')">White paper</button><button onclick="setPageColor('#fffdf4')">Warm paper</button><button onclick="setPageColor('#f3f7ff')">Blue paper</button></div>
  <div class="pcard" id="objectInspector"><h3>Object editor</h3><p id="objectHint">Select a shape or picture on the page to edit it.</p><div id="objectControls"></div></div>
 </aside>
</div>
<button class="sidebar-reopen" id="reopenLeft" onclick="toggleSidebar('left')">☰ Show left panel</button><button class="sidebar-reopen" id="reopenRight" onclick="toggleSidebar('right')">Show right panel ☰</button>
<footer class="status">
 <span id="wordStatus">Words: 0</span><span id="charStatus">Characters: 0</span><span id="pageStatus">Pages: 1</span><span>Language: English</span><span id="saveStatus">Autosave ready</span>
 <span class="push"></span><button onclick="setZoom(-10)">−</button><span id="zoomValue">100%</span><button onclick="setZoom(10)">＋</button>
</footer>

<section class="dashboard" id="dashboard">
 <div class="dash-head"><div><h1>KMS WORD AI Dashboard</h1><p style="color:var(--muted)">A complete workspace for documents, notes and ideas.</p></div><button class="primary" onclick="showDashboard(false)">Open editor →</button></div>
 <div class="dash-grid">
  <div class="dash-card"><div class="emoji">📝</div><h3>Blank document</h3><p>Create a new document and start writing.</p><button class="primary" onclick="createBlank()">Create new</button></div>
  <div class="dash-card"><div class="emoji">📂</div><h3>Open document</h3><p>Open an HTML or TXT file from your computer.</p><button onclick="document.getElementById('openFile').click()">Browse files</button></div>
  <div class="dash-card"><div class="emoji">📄</div><h3>Template library</h3><p>12 editable templates for reports, study, business and events.</p><button onclick="showTemplates()">Choose template</button></div>
   <div class="dash-card"><div class="emoji">🎨</div><h3>Theme gallery</h3><p>12 coordinated palettes for a consistent document design.</p><button onclick="showThemeGallery()">Explore themes</button></div>
  <div class="dash-card"><div class="emoji">📊</div><h3>Document insights</h3><p>Review words, characters, paragraphs and reading time.</p><button onclick="showStatistics()">View insights</button></div>
  <div class="dash-card"><div class="emoji">🎨</div><h3>Whiteboard</h3><p>Sketch concepts, diagrams and quick ideas.</p><button onclick="openWhiteboard()">Open board</button></div>
  <div class="dash-card"><div class="emoji">🖨️</div><h3>Print / PDF</h3><p>Print your document or save it as a PDF.</p><button onclick="requestPrint()">Print document</button></div>
 </div>
 <div class="recent"><h3>Recent documents on this device</h3><div id="recentList">No saved documents yet.</div></div>
</section>

<input id="openFile" type="file" accept=".html,.htm,.txt" hidden onchange="openLocalFile(event)">
<input id="imageFile" type="file" accept="image/*" hidden onchange="insertImageFile(event)">
<div class="overlay" id="overlay"><div class="dialog" id="dialog"></div></div>
<div id="whiteboard"><div class="boardbar"><b style="padding:0 8px">KMS WHITEBOARD PRO</b>
 <button onclick="setBoardTool('pen')">✎ Pen</button><button onclick="setBoardTool('highlighter')">▰ Highlighter</button><button onclick="setBoardTool('eraser')">⌫ Eraser</button><button onclick="setBoardTool('line')">／ Line</button><button onclick="setBoardTool('arrow')">➜ Arrow</button><button onclick="setBoardTool('rect')">□ Rectangle</button><button onclick="setBoardTool('ellipse')">◯ Ellipse</button><button onclick="setBoardTool('triangle')">△ Triangle</button><button onclick="setBoardTool('diamond')">◇ Diamond</button><button onclick="setBoardTool('star')">★ Star</button><button onclick="setBoardTool('text')">T Text</button><button onclick="addStickyNote()">▤ Sticky</button>
 <input type="color" id="boardColor" value="#245bb5" title="Stroke/fill colour"><input type="range" id="boardWidth" min="1" max="40" value="4" title="Stroke width"><select id="boardBg" onchange="setBoardBackground(this.value)"><option value="#ffffff">White</option><option value="#f8fafc">Soft white</option><option value="#fff9e8">Warm</option><option value="#eaf3ff">Blue</option><option value="#18243a">Dark</option></select>
 <button onclick="boardUndo()">↶ Undo</button><button onclick="boardRedo()">↷ Redo</button><button onclick="toggleBoardGrid()">▦ Grid</button><button onclick="boardZoom(-10)">−</button><button onclick="boardZoom(10)">＋</button><button onclick="boardFit()">Fit</button><button onclick="chooseBoardImage()">▧ Image</button><button onclick="clearBoard()">Clear</button><button onclick="downloadBoard()">PNG</button><button onclick="downloadBoardSVG()">SVG</button><button onclick="downloadBoardJPG()">JPG</button><button onclick="downloadBoardJSON()">JSON</button><button onclick="copyBoard()">Copy</button><button onclick="saveBoardProject()">Save project</button><button onclick="loadBoardProject()">Open project</button><button onclick="toggleBoardRuler()">Ruler</button><button onclick="toggleBoardDots()">Dots</button><button onclick="toggleBoardSnap()">Snap</button><button onclick="toggleBoardGuides()">Guides</button><button onclick="toggleBoardCrosshair()">Cursor</button><button onclick="saveBoardLocal()">Save local</button><button onclick="restoreBoardLocal()">Restore local</button><button onclick="mirrorBoard()">Mirror</button><button onclick="rotateBoard90()">Rotate 90°</button><button onclick="invertBoard()">Invert</button><button onclick="addBoardFrame()">Frame</button><button onclick="newBoard()">New board</button><button onclick="closeWhiteboard()">Close ✕</button></div><div class="boardbar" style="background:#203b60;min-height:38px"><b>Palette</b><button onclick="setBoardColor('#172b4d')" style="background:#172b4d;color:white">Navy</button><button onclick="setBoardColor('#245bb5')" style="background:#245bb5;color:white">Blue</button><button onclick="setBoardColor('#16834c')" style="background:#16834c;color:white">Green</button><button onclick="setBoardColor('#c24f78')" style="background:#c24f78;color:white">Rose</button><button onclick="setBoardColor('#ed8b2f')" style="background:#ed8b2f;color:white">Orange</button><button onclick="setBoardColor('#7652b8')" style="background:#7652b8;color:white">Purple</button><button onclick="setBoardColor('#111827')" style="background:#111827;color:white">Black</button><button onclick="setBoardColor('#64748b')" style="background:#64748b;color:white">Grey</button><button onclick="setBoardColor('#0891b2')" style="background:#0891b2;color:white">Cyan</button><button onclick="setBoardColor('#eab308')" style="background:#eab308;color:#111">Yellow</button><b>Stroke</b><button onclick="setBoardWidth(1)">1</button><button onclick="setBoardWidth(3)">3</button><button onclick="setBoardWidth(6)">6</button><button onclick="setBoardWidth(12)">12</button><button onclick="setBoardWidth(20)">20</button><button onclick="setBoardWidth(30)">30</button><b>Background</b><button onclick="setBoardBackground('#ffffff')">Paper</button><button onclick="setBoardBackground('#fff9e8')">Cream</button><button onclick="setBoardBackground('#eaf3ff')">Sky</button><button onclick="setBoardBackground('#f0fdf4')">Mint</button><button onclick="setBoardBackground('#18243a')">Night</button></div>
 <div class="boardstage"><canvas id="board"></canvas></div><div class="boardstatus"><span id="boardToolStatus">Tool: Pen</span><span id="boardZoomStatus">Zoom: 100%</span><span>60+ drawing, editing and canvas controls</span></div></div><input id="boardImageFile" type="file" accept="image/*" hidden onchange="insertBoardImage(event)">
<div class="toast" id="toast"></div>

<script>
// Early tab fallback: remains available even if a later feature script fails.
(function(){
 const names={home:'Home tools',insert:'Insert tools',layout:'Page layout tools',design:'Design tools',references:'Reference tools',review:'Review tools',view:'View tools',help:'Help and support',filetools:'File tools',draw:'Drawing tools',ailab:'AI Lab',mailings:'Mail merge tools',forms:'Forms tools',data:'Document data tools',security:'Privacy tools',accessibility:'Accessibility tools',developer:'Developer tools',automate:'Automation tools'};
 window.setTab=function(id){document.querySelectorAll('.tab[data-tab]').forEach(function(x){x.classList.toggle('active',x.dataset.tab===id)});const r=document.getElementById('ribbon');if(r)r.textContent=names[id]||'Tools'};
 document.addEventListener('click',function(ev){
  const tab=ev.target.closest && ev.target.closest('.tab[data-tab]');
  if(!tab)return;
  const id=tab.dataset.tab;
  setTimeout(function(){
   const ribbon=document.getElementById('ribbon');
   if(ribbon && !ribbon.innerHTML.trim()) ribbon.textContent=names[id]||'Tools';
  },0);
 });
})();
</script>
<script src="qrc:///qtwebchannel/qwebchannel.js"></script>
<script>
// The web-channel bridge is optional. A bridge failure must never stop the editor UI.
try {
 if(typeof qt!=="undefined" && qt.webChannelTransport && typeof QWebChannel==="function") {
  new QWebChannel(qt.webChannelTransport, channel=>{ window.kmsBridge=channel.objects.kmsBridge; });
 }
} catch(bridgeError) { console.warn('Native bridge unavailable; using browser fallbacks.', bridgeError); }
const page=document.getElementById('page'),ribbon=document.getElementById('ribbon');
window.addEventListener('error',e=>{console.error(e.message,e.filename,e.lineno);const t=document.getElementById('toast');if(t){t.textContent='App error: '+e.message;t.style.display='block'}const r=document.getElementById('ribbon');if(r&&!r.innerHTML.trim())r.innerHTML='<div class=\"rgroup\">Ribbon error — use the tab names above and restart the app if this persists.</div>'});
window.addEventListener('unhandledrejection',e=>console.error('Async error:',e.reason));
let activeTab='home',zoom=100,saveTimer=null,boardTool='pen',drawing=false,lastX=0,lastY=0;
const STORAGE_KEY='kms_word_ai_v19_documents';
const ACTIVE_KEY='kms_word_ai_v19_active';
function escapeHTML(s){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function toast(msg){const el=document.getElementById('toast');el.textContent=msg;el.style.display='block';clearTimeout(el._t);el._t=setTimeout(()=>el.style.display='none',2400)}
function group(label,html){return `<div class="rgroup"><div><div class="rgroup-inner">${html}</div><div class="rlabel">${label}</div></div></div>`}
function tool(label,icon,fn){return `<div class="rtool"><button title="${label}" onclick="${fn}">${icon}</button><small>${label}</small></div>`}
function setTab(tab){
 activeTab=tab;document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('active',x.dataset.tab===tab));
 let h='';
 if(tab==='home'){
  h+=group('Clipboard',tool('Paste','▣','pasteFromClipboard()')+tool('Undo','↶',"doFormat('undo')")+tool('Redo','↷',"doFormat('redo')"));
  h+=group('Font',`<div style="display:flex;gap:4px;flex-wrap:wrap;max-width:300px"><select class="font-name" title="370+ font families" onchange="setFont(this.value)">${FONT_FAMILIES.map(f=>`<option>${f}</option>`).join("")}</select><select class="font-size" onchange="setFontSize(this.value)">${[8,9,10,11,12,14,16,18,20,24,28,32,36,48,72].map(x=>`<option ${x===11?'selected':''}>${x}</option>`).join('')}</select><button onclick="doFormat('bold')"><b>B</b></button><button onclick="doFormat('italic')"><i>I</i></button><button onclick="doFormat('underline')"><u>U</u></button><button onclick="doFormat('strikeThrough')"><s>S</s></button><button onclick="doFormat('superscript')">x²</button><button onclick="doFormat('subscript')">x₂</button><input type="color" title="Font colour" onchange="foreColor(this.value)"><input type="color" value="#fff19a" title="Highlight" onchange="highlight(this.value)"></div>`);
  h+=group('Paragraph',tool('Left','☰',"doFormat('justifyLeft')")+tool('Centre','≡',"doFormat('justifyCenter')")+tool('Right','☷',"doFormat('justifyRight')")+tool('Justify','▤',"doFormat('justifyFull')")+tool('Bullets','•',"doFormat('insertUnorderedList')")+tool('Numbering','1.',"doFormat('insertOrderedList')")+tool('Indent','⇥',"doFormat('indent')")+tool('Outdent','⇤',"doFormat('outdent')"));
  h+=group('Styles',`<select onchange="applyBlock(this.value)"><option value="p">Normal</option><option value="h1">Heading 1</option><option value="h2">Heading 2</option><option value="h3">Heading 3</option><option value="blockquote">Quote</option><option value="pre">Code</option></select>`+tool('250+ Text Styles','Aa','showTextStyles()'));
  h+=group('Lists',tool('20 Bullet Styles','☷','showBulletGallery()')+tool('Numbering','1.',"doFormat('insertOrderedList')"));
  h+=group('Editing',tool('Find','⌕','openFind()')+tool('Replace','⇄','openFind()')+tool('Select all','▣',"doFormat('selectAll')"));
 }else if(tab==='insert'){
  h+=group('Pages',tool('Page break','↵','insertPageBreak()')+tool('Date','▦','insertDate()')+tool('Horizontal line','―',"insertHTML('<hr>')"));
  h+=group('Tables',tool('Table','▦','insertTable()')+tool('Blue table','▦',"styleTable('blue')")+tool('Green table','▦',"styleTable('green')")+tool('Dark table','▦',"styleTable('dark')")+tool('Minimal table','▤',"styleTable('minimal')"));
  h+=group('Media',tool('Picture','▧','chooseImage()')+tool('100+ Shapes','◇','showShapeGallery()')+tool('Whiteboard','✎','openWhiteboard()'));
  h+=group('Links',tool('Hyperlink','🔗','insertLink()')+tool('Bookmark','⚑','insertBookmark()'));
  h+=group('Text',tool('Text box','▤','insertTextBox()')+tool('Symbol','Ω','insertSymbol()')+tool('Equation','∑','insertEquation()')+tool('Callout','❝','insertCallout()')+tool('Checklist','☑','insertChecklist()')+tool('Signature','✍','insertSignatureLine()'));
 }else if(tab==='layout'){
  h+=group('Page setup',tool('Margins','↔','setMargins()')+tool('Orientation','▭','toggleOrientation()')+tool('Page size','▤','setPageSize()')+tool('Header','↥','editHeader()')+tool('Footer','↧','editFooter()')+tool('Page border','▣','designPageBorder()'));
  h+=group('Paragraph',tool('Line spacing','↕','setLineSpacing()')+tool('Columns','▥','setPageColumns()')+tool('Single column','▤','clearPageColumns()')+tool('Spacing','↕','setParagraphSpacing()'));
 }else if(tab==='design'){
  h+=group('Themes',tool('Theme gallery','🎨','showThemeGallery()')+tool('Blue','🔵',"applyTheme('blue')")+tool('Green','🟢',"applyTheme('green')")+tool('Bharat','🇮🇳',"applyTheme('bharat')"));
  h+=group('Page background',tool('White','□',"setPageColor('#ffffff')")+tool('Ivory','▧',"setPageColor('#fffdf4')")+tool('Blue tint','▨',"setPageColor('#f3f7ff')")+tool('Page border','▣','designPageBorder()')+tool('Page designs','✧','pageDesignGallery()')+tool('Watermark','◈','insertWatermark()'));
   h+=group('Theme tools',tool('Theme gallery','🎨','showThemeGallery()')+tool('Cover page','▤','insertCoverPage()'));
 }else if(tab==='references'){
  h+=group('References',tool('Contents','☷','insertTOC()')+tool('Footnote','¹','insertFootnote()')+tool('Citation','❝','insertCitation()')+tool('Source list','📚','insertBibliography()'));
 }else if(tab==='review'){
  h+=group('Proofing',tool('Spell check','✓','spellCheck()')+tool('Word count','≡','showStatistics()')+tool('Find','⌕','openFind()'));
  h+=group('Comments',tool('Add comment','💬','addComment()')+tool('Highlight','🖍',"highlight('#fff19a')"));
 }else if(tab==='view'){
  h+=group('Views',tool('Focus mode','⛶','toggleFocus()')+tool('Dashboard','⌂','showDashboard(true)')+tool('Outline','☷','showOutline()')+tool('Whiteboard','✎','openWhiteboard()'));
  h+=group('Zoom',tool('Zoom out','−','setZoom(-10)')+tool('Zoom in','＋','setZoom(10)')+tool('Reset','1:1','resetZoom()'));
 }else if(['filetools','draw','ailab','mailings','forms','data','security','accessibility','developer','automate'].includes(tab)){
  h+=proRibbonGroup(tab);
 }else{
  h+=group('Help',tool('Quick guide','?','showGuide()')+tool('Shortcuts','⌨','showShortcuts()')+tool('About','ⓘ','aboutApp()'));
 }
 ribbon.innerHTML=h;
}
const FONT_FAMILIES=["Arial", "Arial Black", "Arial Narrow", "Arial Rounded MT Bold", "Aptos", "Aptos Display", "Bahnschrift", "Baskerville", "Bodoni MT", "Book Antiqua", "Bookman Old Style", "Calibri", "Calibri Light", "Cambria", "Cambria Math", "Candara", "Century", "Century Gothic", "Century Schoolbook", "Comic Sans MS", "Consolas", "Constantia", "Corbel", "Courier New", "Cursive", "Didot", "Dubai", "Ebrima", "Franklin Gothic Medium", "Freestyle Script", "French Script MT", "Gabriola", "Garamond", "Georgia", "Gill Sans MT", "Gill Sans", "Goudy Old Style", "Haettenschweiler", "Helvetica", "Helvetica Neue", "High Tower Text", "Impact", "Ink Free", "Javanese Text", "Kalinga", "Kartika", "Lato", "Leelawadee UI", "Lucida Bright", "Lucida Calligraphy", "Lucida Console", "Lucida Fax", "Lucida Handwriting", "Lucida Sans", "Lucida Sans Typewriter", "Lucida Sans Unicode", "Maiandra GD", "Malgun Gothic", "Mangal", "Marlett", "Microsoft Himalaya", "Microsoft JhengHei", "Microsoft New Tai Lue", "Microsoft PhagsPa", "Microsoft Sans Serif", "Microsoft Tai Le", "Microsoft YaHei", "Microsoft Yi Baiti", "MingLiU", "Mistral", "Modern No. 20", "Mongolian Baiti", "Monotype Corsiva", "MS Gothic", "MS Mincho", "MS PGothic", "MS PMincho", "MV Boli", "Myanmar Text", "Nirmala UI", "Niagara Engraved", "Niagara Solid", "Noto Sans", "Noto Serif", "OCR A Extended", "Old English Text MT", "Palatino Linotype", "Papyrus", "Perpetua", "Playbill", "Poor Richard", "Pristina", "Rage Italic", "Ravie", "Rockwell", "Rockwell Extra Bold", "Segoe Print", "Segoe Script", "Segoe UI", "Segoe UI Black", "Segoe UI Emoji", "Segoe UI Historic", "Segoe UI Light", "Segoe UI Semibold", "Segoe UI Symbol", "Showcard Gothic", "SimHei", "SimSun", "Sitka Text", "Snap ITC", "Stencil", "Sylfaen", "Symbol", "Tahoma", "Tempus Sans ITC", "Times New Roman", "Trebuchet MS", "Tw Cen MT", "Verdana", "Viner Hand ITC", "Vivaldi", "Vladimir Script", "Wide Latin", "Yu Gothic", "Agency FB", "Algerian", "Andalus", "Angsana New", "Aparajita", "Arabic Typesetting", "Arial Unicode MS", "Bauhaus 93", "Bell MT", "Berlin Sans FB", "Bernard MT Condensed", "Blackadder ITC", "Bodoni MT Black", "Bradley Hand ITC", "Britannic Bold", "Broadway", "Brush Script MT", "Calisto MT", "Castellar", "Centaur", "Chiller", "Colonna MT", "Cooper Black", "Copperplate Gothic Bold", "Curlz MT", "Elephant", "Eras Bold ITC", "Eras Demi ITC", "Eras Light ITC", "Felix Titling", "Footlight MT Light", "Forte", "Gigi", "Gloucester MT Extra Condensed", "Goudy Stout", "Jokerman", "Juice ITC", "Kristen ITC", "Magneto", "Niagara Engraved", "Onyx", "Perpetua Titling MT", "Script MT Bold", "Tw Cen MT Condensed", 'Abadi MT Condensed Light', 'Aharoni', 'Aldhabi', 'Algerian', 'American Typewriter', 'Andale Mono', 'Aptos Narrow', 'Arial Nova', 'Arial Nova Cond', 'Arial Nova Light', 'Arial Nova', 'Athelas', 'Avenir', 'Avenir Next', 'Baskerville Old Face', 'Bauhaus', 'Bebas Neue', 'Bell MT', 'Bodoni 72', 'Bodoni 72 Smallcaps', 'Bradley Hand', 'Calibri Math', 'Calisto MT', 'Charter', 'Chalkboard', 'Chalkboard SE', 'Chalkduster', 'Cochin', 'Comic Sans', 'Copperplate', 'Courier', 'Courier Prime', 'DIN Alternate', 'DIN Condensed', 'Euphemia UCAS', 'Futura', 'Geneva', 'Georgia Pro', 'Gill Sans Nova', 'Helvetica Neue', 'Herculanum', 'Hoefler Text', 'Iowan Old Style', 'Kefa', 'Luminari', 'Menlo', 'Monaco', 'Noteworthy', 'Optima', 'Palatino', 'Papyrus', 'Party LET', 'Phosphate', 'Rockwell Nova', 'Savoye LET', 'SignPainter', 'Skia', 'Snell Roundhand', 'Superclarendon', 'Trattatello', 'Verdana Pro', 'Aptos Mono', 'Aptos Serif', 'Aptos Display', 'Aptos Light', 'Aptos Black', 'Bahnschrift SemiCondensed', 'Bahnschrift Light', 'Bahnschrift SemiBold', 'Bahnschrift Condensed', 'Segoe UI Variable', 'Segoe UI Variable Display', 'Segoe UI Variable Text', 'Segoe UI Emoji', 'Noto Sans Devanagari', 'Noto Serif Devanagari', 'Noto Sans Bengali', 'Noto Serif Bengali', 'Noto Sans Tamil', 'Noto Sans Telugu', 'Noto Sans Gujarati', 'Noto Sans Gurmukhi', 'Noto Sans Arabic', 'Noto Naskh Arabic', 'Noto Sans Hebrew', 'Noto Sans Thai', 'Noto Sans CJK SC', 'Noto Sans CJK JP', 'Noto Sans CJK KR', 'Liberation Sans', 'Liberation Serif', 'Liberation Mono', 'Carlito', 'Cousine', 'Arimo', 'Tinos', 'Ubuntu', 'Ubuntu Mono', 'Ubuntu Condensed', 'Cantarell', 'DejaVu Sans', 'DejaVu Serif', 'DejaVu Sans Mono'];

function focusPage(){page.focus()}
function doFormat(cmd){focusPage();document.execCommand(cmd,false,null);updateStats()}
function setFont(v){focusPage();document.execCommand('fontName',false,v)}
function setFontSize(v){focusPage();let m={8:1,9:2,10:2,11:3,12:3,14:4,16:4,18:5,20:5,24:5,28:6,32:6,36:6,48:7,72:7};document.execCommand('fontSize',false,m[v]||3)}
function foreColor(v){focusPage();document.execCommand('foreColor',false,v)}
function highlight(v){focusPage();document.execCommand('hiliteColor',false,v)}
function applyBlock(v){focusPage();document.execCommand('formatBlock',false,v)}
function insertHTML(s){focusPage();document.execCommand('insertHTML',false,s);updateStats()}
function insertTable(){let r=prompt('Number of rows (1–25):','3');if(!r)return;let c=prompt('Number of columns (1–15):','3');if(!c)return;r=Math.max(1,Math.min(25,parseInt(r)||3));c=Math.max(1,Math.min(15,parseInt(c)||3));let s='<table><tbody>'+Array.from({length:r},(_,i)=>'<tr>'+Array.from({length:c},()=>i===0?'<th>Heading</th>':'<td>Type here</td>').join('')+'</tr>').join('')+'</tbody></table><p></p>';insertHTML(s)}
function chooseImage(){document.getElementById('imageFile').click()}
function insertImageFile(e){let f=e.target.files[0];if(!f)return;let rd=new FileReader();rd.onload=()=>insertHTML(`<img class="kms-image" src="${rd.result}" alt="Inserted image" style="width:320px;height:220px;object-fit:contain"><p></p>`);rd.readAsDataURL(f);e.target.value=''}
function insertLink(){let u=prompt('Enter a full website address:','https://');if(u)insertHTML(`<a href="${escapeHTML(u)}" target="_blank">${escapeHTML(u)}</a>`)}
function insertBookmark(){let n=prompt('Bookmark label:','Section');if(n)insertHTML(`<a id="${escapeHTML(n.replace(/[^a-z0-9_-]/gi,'_'))}"></a>`)}
function insertShape(){showShapeGallery()}
function insertTextBox(){insertHTML('<div style="border:1px solid #245bb5;padding:14px;margin:10px 0;background:#f4f7ff">Text box — click here to type</div>')}
function insertSymbol(){let s=prompt('Enter or paste a symbol:','₹  ©  ®  ™  Ω  ∑  ✓  ∞');if(s)insertHTML(escapeHTML(s))}
function insertEquation(){let s=prompt('Enter an equation or formula as text:','a² + b² = c²');if(s)insertHTML(`<span style="font-family:Georgia;font-size:1.15em">${escapeHTML(s)}</span>`)}
function insertDate(){insertHTML(new Date().toLocaleDateString('en-IN',{day:'2-digit',month:'long',year:'numeric'}))}
function insertPageBreak(){insertHTML('<div style="page-break-before:always;border-top:1px dashed #aaa;margin:25px 0;padding-top:8px;color:#888">— Page break —</div>')}
function insertFootnote(){let s=prompt('Footnote text:');if(s)insertHTML(`<sup>[1]</sup><small> [Footnote: ${escapeHTML(s)}]</small>`)}
function insertCitation(){let s=prompt('Citation / source details:');if(s)insertHTML(`<span style="color:#245bb5">(${escapeHTML(s)})</span>`)}
function insertBibliography(){insertHTML('<h2>References / Bibliography</h2><ol><li>Enter source details here.</li></ol>')}
function insertTOC(){let hs=[...page.querySelectorAll('h1,h2,h3')].map((h,i)=>`<li>${escapeHTML(h.innerText)}</li>`).join('');insertHTML(`<h2>Table of Contents</h2><ol>${hs||'<li>Add headings to your document first.</li>'}</ol>`)}
function insertColumns(){insertHTML('<table style="table-layout:fixed"><tr><td style="width:50%;vertical-align:top">Column 1</td><td style="width:50%;vertical-align:top">Column 2</td></tr></table>')}
function setMargins(){let n=prompt('Page padding in pixels (30–140):','80');if(n!==null){n=Math.max(30,Math.min(140,parseInt(n)||80));page.style.padding=n+'px'}}
function toggleOrientation(){let wide=page.dataset.landscape!=='1';page.dataset.landscape=wide?'1':'0';let w=wide?1056:816;page.style.width=w+'px';document.querySelector('.page-shell').style.width=w+'px';toast(wide?'Landscape layout':'Portrait layout')}
const PAGE_SIZES=[["A4", 794, 1123], ["A3", 1123, 1587], ["A5", 559, 794], ["A6", 397, 559], ["A2", 1587, 2245], ["A1", 2245, 3179], ["A0", 3179, 4494], ["Letter", 816, 1056], ["Legal", 816, 1344], ["Tabloid", 1056, 1632], ["Executive", 696, 1008], ["Statement", 528, 816], ["Folio", 816, 1248], ["Ledger", 1632, 1056], ["B4", 945, 1334], ["B5", 665, 945], ["B6", 472, 665], ["B3", 1334, 1889], ["B2", 1889, 2672], ["B1", 2672, 3780], ["B0", 3780, 5346], ["C4", 918, 1298], ["C5", 649, 918], ["C6", 459, 649], ["DL Envelope", 417, 827], ["#10 Envelope", 396, 918], ["Monarch Envelope", 504, 972], ["Photo 4x6", 384, 576], ["Photo 5x7", 480, 672], ["Photo 8x10", 768, 960], ["Square 8x8", 768, 768], ["Square 10x10", 960, 960], ["Wide 16:9", 1120, 630], ["Presentation 4:3", 960, 720], ["Web page", 960, 1200], ["Index Card", 576, 384], ["Postcard", 576, 384], ["Compact", 680, 900], ["Standard", 816, 1056], ["Large", 900, 1200]];

function setPageSize(){
 const opts=PAGE_SIZES.map((x,i)=>`<option value="${i}" ${x[0]==='A4'?'selected':''}>${x[0]} — ${x[1]} × ${x[2]} px</option>`).join('');
 showDialog(`<h2>Page size library · 40 formats</h2><label>Choose page format</label><select id="pageSizeChoice" size="12" style="height:260px">${opts}</select><div class="dialog-actions"><button onclick="closeDialog()">Cancel</button><button class="primary" onclick="applyPageSize()">Apply size</button></div>`);
}
function applyPageSize(){const x=PAGE_SIZES[Number(document.getElementById('pageSizeChoice').value)];if(!x)return;page.style.width=x[1]+'px';page.style.minHeight=x[2]+'px';document.querySelector('.page-shell').style.width=x[1]+'px';page.dataset.pageSize=x[0];page.dataset.pageWidth=x[1];page.dataset.pageHeight=x[2];closeDialog();toast(x[0]+' page size applied');updateStats()}
function toggleBorder(){designPageBorder()}
const BORDER_PRESETS=Array.from({length:100},(_,i)=>{const colors=['#294d7a','#245bb5','#16834c','#7652b8','#e27e20','#b58a36','#c24f78','#087e83','#34445a','#8f2636'];const styles=['solid','double','dashed','dotted','groove','ridge','inset','outset'];return {name:'Design '+String(i+1).padStart(2,'0'),color:colors[i%10],style:styles[i%8],width:[1,2,3,4,5,6,7,8][i%8],radius:(i%6)*3}});
function designPageBorder(){const cards=BORDER_PRESETS.map((b,i)=>`<button class="gallery-item" onclick="applyBorderPreset(${i})"><span style="display:block;width:76px;height:50px;background:#fff;border:${Math.min(b.width,6)}px ${b.style} ${b.color};border-radius:${b.radius}px"></span><small>${b.name}</small></button>`).join('');showDialog(`<h2>100 page border designs</h2><p>Preview and apply a virtual border design.</p><div class="gallery-grid">${cards}</div><div class="dialog-actions"><button onclick="closeDialog()">Close</button></div>`)}
function applyBorderPreset(i){const b=BORDER_PRESETS[i];page.style.border=`${b.width}px ${b.style} ${b.color}`;page.style.borderRadius=b.radius+'px';page.dataset.borderPreset=b.name;closeDialog();toast(b.name+' applied')}
function applyPageBorder(){let style=document.getElementById('borderStyle').value,color=document.getElementById('borderColor').value,width=document.getElementById('borderWidth').value;page.className='page';if(style==='double'){page.classList.add('page-border-double');page.style.borderColor=color;page.style.borderWidth=(width*2)+'px'}else if(style==='shadow'){page.classList.add('page-border-shadow');page.style.borderColor=color}else if(style==='dashed'){page.classList.add('page-border-dashed');page.style.borderColor=color;page.style.borderWidth=width+'px'}else if(style==='rounded'){page.classList.add('page-border-rounded');page.style.borderColor=color;page.style.borderWidth=width+'px'}else if(style==='classic'){page.classList.add('page-border-classic');page.style.borderColor=color}else if(style==='gold'){page.classList.add('page-border-gold');page.style.borderColor=color}else if(style==='solid'){page.style.border=width+'px solid '+color}else{page.style.border='';page.style.outline='';page.style.boxShadow='var(--shadow)'}closeDialog();toast('Page border applied')}
function pageDesignGallery(){showDialog(`<h2>Page design gallery</h2><p>Apply a coordinated document look.</p><div class="dash-grid"><div class="dash-card"><h3>Professional blue</h3><p>Blue headings and clean blue frame.</p><button onclick="applyPageDesign('professional')">Apply</button></div><div class="dash-card"><h3>Academic</h3><p>Serif typography and classic double border.</p><button onclick="applyPageDesign('academic')">Apply</button></div><div class="dash-card"><h3>Modern minimal</h3><p>Subtle grey lines and spacious layout.</p><button onclick="applyPageDesign('modern')">Apply</button></div><div class="dash-card"><h3>Executive gold</h3><p>Elegant gold frame and formal headings.</p><button onclick="applyPageDesign('gold')">Apply</button></div><div class="dash-card"><h3>Creative teal</h3><p>Teal headings and soft tinted page.</p><button onclick="applyPageDesign('creative')">Apply</button></div><div class="dash-card"><h3>Classic report</h3><p>Traditional black headings and inset frame.</p><button onclick="applyPageDesign('classic')">Apply</button></div></div>`)}
function applyPageDesign(name){let d={professional:['#174b91','#ffffff','solid','#174b91'],academic:['#303846','#fffefa','double','#303846'],modern:['#34495e','#ffffff','shadow','#9aa9bd'],gold:['#8a6728','#fffdf7','gold','#b58a36'],creative:['#087e83','#f1fbfa','rounded','#087e83'],classic:['#222222','#ffffff','classic','#555555']}[name];page.style.background=d[1];page.querySelectorAll('h1,h2,h3').forEach(h=>h.style.color=d[0]);page.className='page';let st=d[2];if(st==='double')page.classList.add('page-border-double');else if(st==='shadow')page.classList.add('page-border-shadow');else if(st==='gold')page.classList.add('page-border-gold');else if(st==='rounded')page.classList.add('page-border-rounded');else if(st==='classic')page.classList.add('page-border-classic');else page.style.border='2px solid '+d[3];page.style.borderColor=d[3];closeDialog();toast('Page design applied')}
function addPageField(kind){let cls=kind==='header'?'page-header':'page-footer',existing=page.querySelector('.'+cls);if(existing){existing.focus();toast(kind+' selected — edit directly');return}let el=document.createElement('div');el.className=cls;el.contentEditable='true';el.setAttribute('data-page-field',kind);el.innerHTML=kind==='header'?'Document title · Organisation':'<span>Confidential</span><span style="float:right">Page 1</span>';if(kind==='header')page.insertBefore(el,page.firstChild);else page.appendChild(el);el.focus();updateStats();toast(kind+' added — click text to edit')}
function editHeader(){addPageField('header')}
function editFooter(){addPageField('footer')}
function stylePageField(kind,align){let el=page.querySelector('.page-'+kind);if(!el){addPageField(kind);el=page.querySelector('.page-'+kind)}el.style.textAlign=align;el.style.fontFamily='Georgia,serif';el.style.letterSpacing='.4px'}
function setLineSpacing(){let v=prompt('Line spacing: 1, 1.15, 1.5 or 2','1.5');if(['1','1.15','1.5','2'].includes(v))page.style.lineHeight=v}
function setPageColor(c){page.style.background=c}
function applyTheme(t){let c={blue:'#174b91',green:'#176b45',classic:'#333333',bharat:'#d66b16'}[t];page.querySelectorAll('h1,h2,h3').forEach(h=>h.style.color=c);toast('Document heading theme applied')}
function openFind(){showDialog(`<h2>Find & Replace</h2><label>Find text</label><input id="findValue" placeholder="Enter text to find"><label>Replace with</label><input id="replaceValue" placeholder="Replacement text"><div class="dialog-actions"><button onclick="closeDialog()">Close</button><button onclick="findNext()">Find next</button><button class="primary" onclick="replaceAll()">Replace all</button></div>`)}
function findNext(){let q=document.getElementById('findValue').value;if(q&&!window.find(q))toast('No further match found')}
function replaceAll(){let q=document.getElementById('findValue').value,r=document.getElementById('replaceValue').value;if(!q)return;let w=document.createTreeWalker(page,NodeFilter.SHOW_TEXT),a=[];while(w.nextNode())a.push(w.currentNode);let count=0;a.forEach(n=>{if(n.parentElement.closest('script,style'))return;let parts=n.nodeValue.split(q);count+=parts.length-1;if(parts.length>1)n.nodeValue=parts.join(r)});closeDialog();updateStats();toast(count+' replacement(s) made')}
function spellCheck(){page.setAttribute('spellcheck','true');page.focus();toast('Browser spell-check is enabled. Check your system language settings.')}
function addComment(){let c=prompt('Comment:');if(c)insertHTML(`<aside style="border-left:3px solid #e28b2f;background:#fff8e9;padding:9px;margin:9px 0;color:#75501d"><b>Comment:</b> ${escapeHTML(c)}</aside>`)}
function showStatistics(){let t=page.innerText.trim(),w=t?t.split(/\s+/).length:0,ch=t.length,paras=t?t.split(/\n+/).filter(Boolean).length:0,heads=page.querySelectorAll('h1,h2,h3').length,read=Math.max(1,Math.ceil(w/200));showDialog(`<h2>Document statistics</h2><div class="dash-grid"><div class="dash-card"><h3>${w}</h3><p>Words</p></div><div class="dash-card"><h3>${ch}</h3><p>Characters</p></div><div class="dash-card"><h3>${paras}</h3><p>Text lines / paragraphs</p></div><div class="dash-card"><h3>${heads}</h3><p>Headings</p></div><div class="dash-card"><h3>${read} min</h3><p>Estimated reading time</p></div></div><div class="dialog-actions"><button class="primary" onclick="closeDialog()">Close</button></div>`)}
function updateOutline(){let hs=[...page.querySelectorAll('h1,h2,h3')],box=document.getElementById('outlineList');box.innerHTML=hs.length?hs.map((h,i)=>`<button style="padding-left:${(parseInt(h.tagName[1])-1)*10+5}px" onclick="jumpToHeading(${i})">${escapeHTML(h.innerText||'Untitled heading')}</button>`).join(''):'<p>No headings yet.</p>'}
function jumpToHeading(i){let hs=[...page.querySelectorAll('h1,h2,h3')];hs[i]?.scrollIntoView({behavior:'smooth',block:'center'})}
function showOutline(){updateOutline();toast('Outline updated in the right panel')}
function updateStats(){let t=page.innerText.trim(),w=t?t.split(/\s+/).length:0,c=t.length;document.getElementById('wordStatus').textContent='Words: '+w;document.getElementById('charStatus').textContent='Characters: '+c;document.getElementById('stats').textContent=`${w} words · ${c} characters`;document.getElementById('pageStatus').textContent='Pages: '+Math.max(1,Math.ceil(page.scrollHeight/1056));updateOutline();scheduleAutosave()}
function docs(){try{return JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}')}catch(e){return {}}}
function saveDocument(){try{let all=docs(),id=localStorage.getItem(ACTIVE_KEY)||'doc_'+Date.now();localStorage.setItem(ACTIVE_KEY,id);all[id]={id,title:document.getElementById('docTitle').value||'Untitled Document',html:page.innerHTML,updated:new Date().toISOString()};localStorage.setItem(STORAGE_KEY,JSON.stringify(all));document.getElementById('saveStatus').textContent='Saved locally';renderRecent();toast('Document saved on this device')}catch(e){document.getElementById('saveStatus').textContent='Storage unavailable — export a copy';toast('Could not save locally. Use Export HTML.')}}
function scheduleAutosave(){document.getElementById('saveStatus').textContent='Editing…';clearTimeout(saveTimer);saveTimer=setTimeout(()=>{saveDocumentQuiet()},900)}
function saveDocumentQuiet(){try{let all=docs(),id=localStorage.getItem(ACTIVE_KEY)||'doc_'+Date.now();localStorage.setItem(ACTIVE_KEY,id);all[id]={id,title:document.getElementById('docTitle').value||'Untitled Document',html:page.innerHTML,updated:new Date().toISOString()};localStorage.setItem(STORAGE_KEY,JSON.stringify(all));document.getElementById('saveStatus').textContent='Autosaved';renderRecent()}catch(e){document.getElementById('saveStatus').textContent='Storage unavailable — export a copy'}}
function createBlank(){page.innerHTML='<h1>Untitled Document</h1><p>Start writing here…</p>';document.getElementById('docTitle').value='Untitled Document';try{localStorage.setItem(ACTIVE_KEY,'doc_'+Date.now())}catch(e){}showDashboard(false);updateStats();toast('New document created')}
function newDocument(){if(!confirm('Create a new document? Make sure your current work is saved.'))return;createBlank()}
function renderRecent(){let all=Object.values(docs()).sort((a,b)=>b.updated.localeCompare(a.updated)),el=document.getElementById('recentList');if(!all.length){el.textContent='No saved documents yet.';return}el.innerHTML=all.slice(0,12).map(d=>`<div class="recent-row"><span>📄</span><span class="grow"><b>${escapeHTML(d.title)}</b><br><small style="color:#8290a3">${new Date(d.updated).toLocaleString()}</small></span><button onclick="loadDocument('${d.id}')">Open</button><button class="warn" onclick="deleteDocument('${d.id}')">Delete</button></div>`).join('')}
function loadDocument(id){let d=docs()[id];if(!d)return;page.innerHTML=d.html;document.getElementById('docTitle').value=d.title;try{localStorage.setItem(ACTIVE_KEY,id)}catch(e){}showDashboard(false);updateStats();toast('Document opened')}
function deleteDocument(id){if(!confirm('Delete this saved document from this device?'))return;let a=docs();delete a[id];try{localStorage.setItem(STORAGE_KEY,JSON.stringify(a))}catch(e){}renderRecent()}
function showDashboard(show){document.getElementById('dashboard').style.display=show?'block':'none';if(show)renderRecent()}
function exportHTML(){let title=document.getElementById('docTitle').value||'KMS_Document',src='<!doctype html><html><head><meta charset="utf-8"><title>'+escapeHTML(title)+'</title><style>body{font:11pt Arial;line-height:1.5;max-width:850px;margin:40px auto;padding:20px}table{border-collapse:collapse;width:100%}td,th{border:1px solid #777;padding:7px}.page-border-double{border:5px double #294d7a}.page-border-shadow{border:1px solid #9aa9bd;box-shadow:0 0 0 7px #e7edf6}.page-border-dashed{border:3px dashed #567da8}.page-border-rounded{border:4px solid #2b659f;border-radius:18px}.page-border-classic{border:2px solid #333;outline:1px solid #999;outline-offset:-9px}.page-border-gold{border:7px double #b58a36}.page-header{border-bottom:1px solid #9aa9bd;padding-bottom:10px;margin-bottom:24px}.page-footer{border-top:1px solid #9aa9bd;padding-top:10px;margin-top:26px}</style></head><body>'+page.innerHTML+'</body></html>';download(src,slug(title)+'.html','text/html');}
function exportTXT(){download(page.innerText,slug(document.getElementById('docTitle').value||'KMS_Document')+'.txt','text/plain')}
function slug(s){return s.replace(/[^\w-]+/g,'_').slice(0,70)||'KMS_Document'}
function download(data,name,type){let u=URL.createObjectURL(new Blob([data],{type})),a=document.createElement('a');a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),1000);toast('Download started')}
function openLocalFile(e){let f=e.target.files[0];if(!f)return;let rd=new FileReader();rd.onload=()=>{if(/\.txt$/i.test(f.name))page.innerHTML='<pre>'+escapeHTML(rd.result)+'</pre>';else{let d=new DOMParser().parseFromString(rd.result,'text/html');page.innerHTML=d.body.innerHTML}document.getElementById('docTitle').value=f.name.replace(/\.(html?|txt)$/i,'');localStorage.setItem(ACTIVE_KEY,'doc_'+Date.now());showDashboard(false);updateStats();toast('File opened')};rd.readAsText(f);e.target.value=''}
function showTemplates(){
 const items=[
  ['letter','Formal letter','Professional correspondence','✉️'],
  ['report','Project report','Structured report with sections','📊'],
  ['meeting','Meeting minutes','Agenda, decisions and action items','🗒️'],
  ['study','Study notes','Concepts, definitions and revision','📚'],
  ['resume','Modern résumé','Profile, skills and experience','👤'],
  ['invoice','Business invoice','Itemised bill and payment details','🧾'],
  ['proposal','Business proposal','Executive summary and deliverables','💼'],
  ['research','Research paper','Abstract, methodology and references','🔬'],
  ['newsletter','Newsletter','Editorial layout with highlights','📰'],
  ['lesson','Lesson plan','Objectives, activities and assessment','🎓'],
  ['event','Event programme','Schedule, speakers and details','🎤'],
  ['planner','Weekly planner','Priorities, tasks and weekly goals','🗓️']
 ];
 showDialog(`<h2>Template library</h2><p>Choose a professionally structured starting point. You can edit every section after inserting it.</p><div class="dash-grid">${items.map(x=>`<div class="template-card"><div style="font-size:24px">${x[3]}</div><h3>${x[1]}</h3><p>${x[2]}</p><button class="primary" onclick="useTemplate('${x[0]}')">Use template</button></div>`).join('')}</div><div class="dialog-actions"><button onclick="closeDialog()">Close</button></div>`);
}
function useTemplate(t){
 const templates={
 letter:`<p style="text-align:right">[Your address]<br>[City, PIN]</p><p>[Date]</p><p>To,<br>[Recipient name]<br>[Organisation / Address]</p><h2>Subject: [Write subject]</h2><p>Dear Sir/Madam,</p><p>[Opening paragraph]</p><p>[Main message and supporting details]</p><p>[Closing paragraph]</p><p>Yours faithfully,<br><b>[Your name]</b></p>`,
 report:`<h1 style="text-align:center">PROJECT REPORT</h1><p style="text-align:center;color:#64748b">[Project title]</p><p style="text-align:center">Prepared by [Name] · [Date]</p><hr><h2>Executive summary</h2><p>[Summarise the report in a few sentences.]</p><h2>1. Introduction</h2><p>[Background and context]</p><h2>2. Objectives</h2><ul><li>[Objective one]</li><li>[Objective two]</li></ul><h2>3. Methodology</h2><p>[Explain the approach.]</p><h2>4. Findings and discussion</h2><p>[Present key findings.]</p><h2>5. Conclusion</h2><p>[Summarise outcomes and recommendations.]</p><h2>References</h2><ol><li>[Source details]</li></ol>`,
 meeting:`<h1>Meeting Minutes</h1><div class="note-box"><b>Date:</b> [Date] &nbsp; <b>Time:</b> [Time]<br><b>Location:</b> [Location] &nbsp; <b>Chair:</b> [Name]</div><h2>Attendees</h2><ul><li>[Name / Role]</li></ul><h2>Agenda</h2><ol><li>[Agenda item]</li></ol><h2>Discussion and decisions</h2><p>[Record the key discussion.]</p><h2>Action items</h2><table class="table-blue"><tr><th>Action</th><th>Owner</th><th>Due date</th><th>Status</th></tr><tr><td>[Task]</td><td>[Name]</td><td>[Date]</td><td>Open</td></tr></table>`,
 study:`<h1>Study Notes</h1><p><b>Subject:</b> [Subject] &nbsp; <b>Chapter:</b> [Chapter]</p><div class="section-band">Learning objectives</div><ul><li>[What should you understand?]</li></ul><h2>Key concepts</h2><p>[Explain the main ideas.]</p><h2>Definitions</h2><table class="table-green"><tr><th>Term</th><th>Meaning</th></tr><tr><td>[Term]</td><td>[Definition]</td></tr></table><h2>Important formulas</h2><div class="callout">[Formula and explanation]</div><h2>Solved examples</h2><p>[Worked example]</p><h2>Quick revision</h2><ul><li>[Key point]</li></ul>`,
 resume:`<h1 style="text-align:center">[YOUR NAME]</h1><p style="text-align:center">[Phone] · [Email] · [City]</p><hr><h2>Professional profile</h2><p>[A concise summary of your strengths and goals.]</p><h2>Education</h2><table class="table-minimal"><tr><th>Qualification</th><th>Institution</th><th>Year</th></tr><tr><td>[Qualification]</td><td>[Institution]</td><td>[Year]</td></tr></table><h2>Skills</h2><ul><li>[Skill]</li><li>[Skill]</li></ul><h2>Projects and experience</h2><p><b>[Project / Role]</b> — [Description and outcome]</p><h2>Achievements</h2><ul><li>[Achievement]</li></ul><h2>Interests</h2><p>[Relevant interests]</p>`,
 invoice:`<h1 style="text-align:right">INVOICE</h1><p><b>[Business name]</b><br>[Address] · [Email] · [Phone]</p><hr><p><b>Invoice no:</b> [INV-001] &nbsp; <b>Date:</b> [Date]<br><b>Bill to:</b> [Customer name]</p><table class="table-blue"><tr><th>Description</th><th>Qty</th><th>Rate</th><th>Amount</th></tr><tr><td>[Item / service]</td><td>1</td><td>₹0.00</td><td>₹0.00</td></tr><tr><td colspan="3" style="text-align:right"><b>Total</b></td><td><b>₹0.00</b></td></tr></table><p><b>Payment terms:</b> [Terms]</p><p>Thank you for your business.</p>`,
 proposal:`<h1 style="text-align:center">PROJECT PROPOSAL</h1><p style="text-align:center">[Project name] · [Prepared for] · [Date]</p><hr><h2>Executive summary</h2><p>[Proposal overview]</p><h2>Problem / opportunity</h2><p>[Describe the need.]</p><h2>Proposed solution</h2><p>[Describe your solution.]</p><h2>Scope and deliverables</h2><ul><li>[Deliverable]</li></ul><h2>Timeline</h2><table class="table-blue"><tr><th>Phase</th><th>Activity</th><th>Timeline</th></tr><tr><td>1</td><td>[Activity]</td><td>[Duration]</td></tr></table><h2>Budget</h2><p>[Estimated costs]</p><h2>Conclusion</h2><p>[Next steps]</p>`,
 research:`<h1 style="text-align:center">[RESEARCH TITLE]</h1><p style="text-align:center">[Author] · [Institution] · [Date]</p><h2>Abstract</h2><p>[Brief summary of the research.]</p><h2>Keywords</h2><p>[Keyword 1]; [Keyword 2]; [Keyword 3]</p><h2>1. Introduction</h2><p>[Research context and question]</p><h2>2. Literature review</h2><p>[Relevant prior work]</p><h2>3. Methodology</h2><p>[Methods and materials]</p><h2>4. Results and discussion</h2><p>[Findings and interpretation]</p><h2>5. Conclusion</h2><p>[Conclusion and future work]</p><h2>References</h2><ol><li>[Citation]</li></ol>`,
 newsletter:`<p class="cover-kicker">[Organisation] · [Month Year]</p><h1 class="cover-title">THE [NEWSLETTER] UPDATE</h1><p class="cover-subtitle">Ideas, updates and stories from our community</p><hr><h2>In this issue</h2><ul><li>[Top story]</li><li>[Announcement]</li></ul><h2>Featured story</h2><p>[Write your main story here.]</p><h2>News and announcements</h2><p>[Updates]</p><div class="callout"><b>Save the date:</b> [Event details]</div><h2>Coming next</h2><p>[Preview]</p>`,
 lesson:`<h1>Lesson Plan</h1><p><b>Subject:</b> [Subject] &nbsp; <b>Class:</b> [Class] &nbsp; <b>Duration:</b> [Time]</p><h2>Learning objectives</h2><ul><li>Students will be able to [objective].</li></ul><h2>Prerequisite knowledge</h2><p>[Prior learning]</p><h2>Teaching resources</h2><ul><li>[Resource]</li></ul><h2>Lesson sequence</h2><table class="table-green"><tr><th>Time</th><th>Teacher activity</th><th>Student activity</th></tr><tr><td>[Min]</td><td>[Explain]</td><td>[Participate]</td></tr></table><h2>Assessment</h2><p>[How learning will be checked]</p><h2>Homework / follow-up</h2><p>[Task]</p>`,
 event:`<h1 style="text-align:center">[EVENT NAME]</h1><p style="text-align:center;color:#64748b">[Theme or tagline]</p><div class="callout" style="text-align:center"><h2>[Date] · [Time]</h2><p>[Venue / Online details]</p></div><h2>Welcome</h2><p>[Event introduction]</p><h2>Programme schedule</h2><table class="table-blue"><tr><th>Time</th><th>Session</th><th>Speaker</th></tr><tr><td>[Time]</td><td>[Session]</td><td>[Name]</td></tr></table><h2>Organising team</h2><p>[Names and roles]</p><h2>Contact</h2><p>[Contact details]</p>`,
 planner:`<h1>Weekly Planner</h1><p><b>Week of:</b> [Date] &nbsp; <b>Main focus:</b> [Priority]</p><h2>Top priorities</h2><ol><li>[Priority one]</li><li>[Priority two]</li><li>[Priority three]</li></ol><h2>Weekly schedule</h2><table class="table-blue"><tr><th>Day</th><th>Main task</th><th>Other tasks</th><th>Done</th></tr>${['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'].map(d=>`<tr><td>${d}</td><td>[Task]</td><td>[Task]</td><td>☐</td></tr>`).join('')}</table><h2>Notes and reflections</h2><p>[What went well? What needs attention?]</p>`
 };
 page.innerHTML=templates[t]||'<h1>Untitled</h1><p>Start writing here…</p>';
 document.getElementById('docTitle').value=t.charAt(0).toUpperCase()+t.slice(1)+' Document';
 localStorage.setItem(ACTIVE_KEY,'doc_'+Date.now());closeDialog();showDashboard(false);updateStats();toast('Template inserted');
}
function toggleSidebar(which){const cls=which==='left'?'hide-left':'hide-right';document.body.classList.toggle(cls);const hidden=document.body.classList.contains(cls);const el=document.getElementById(which==='left'?'leftSidebar':'rightSidebar');if(el)el.style.display=hidden?'none':'';const btn=document.getElementById(which==='left'?'reopenLeft':'reopenRight');if(btn)btn.style.display=hidden?'block':'none';}
function requestPrint(){try{showDashboard(false);closeDialog();page.focus();if(window.kmsBridge&&typeof window.kmsBridge.printDocument==='function'){window.kmsBridge.printDocument();return}window.focus();window.print();}catch(e){console.error('Print error',e);toast('Print could not start. Try Ctrl+P or File > Print.')}}
function setZoom(n){zoom=Math.max(60,Math.min(160,zoom+n));page.style.zoom=zoom/100;document.getElementById('zoomValue').textContent=zoom+'%'}
function resetZoom(){zoom=100;page.style.zoom=1;document.getElementById('zoomValue').textContent='100%'}
function toggleFocus(){let els=[document.querySelector('.topbar'),document.querySelector('.tabs'),document.querySelector('.ribbon'),document.querySelector('.sidebar'),document.querySelector('.rightpanel')];let hide=els[0].style.display!=='none';els.forEach(x=>x.style.display=hide?'none':'');document.getElementById('workspace').style.height=hide?'calc(100vh - 32px)':'calc(100vh - 238px)'}
function toggleDark(){document.body.classList.toggle('dark')}
function pasteFromClipboard(){navigator.clipboard?.readText().then(t=>insertHTML(escapeHTML(t).replace(/\n/g,'<br>'))).catch(()=>toast('Use Ctrl+V to paste'))}
function runCommandSearch(){let q=document.getElementById('commandSearch').value.toLowerCase().trim(),m={bold:()=>doFormat('bold'),italic:()=>doFormat('italic'),underline:()=>doFormat('underline'),table:insertTable,image:chooseImage,find:openFind,replace:openFind,whiteboard:openWhiteboard,dashboard:()=>showDashboard(true),save:saveDocument,export:exportHTML,print:requestPrint,'word count':showStatistics,template:showTemplates,themes:showThemeGallery,watermark:insertWatermark,'cover page':insertCoverPage,checklist:insertChecklist,callout:insertCallout};if(m[q])m[q]();else toast('Try: bold, table, image, find, replace, whiteboard, save, export, print, themes, watermark, cover page, checklist')}
function showThemeGallery(){
 const themes=[
  ['ocean','Ocean Blue','#174b91','#eaf2ff','#245bb5'],
  ['forest','Forest Green','#176b45','#edf8f1','#16834c'],
  ['royal','Royal Purple','#583c91','#f3effb','#7652b8'],
  ['sunrise','Sunrise Orange','#a94f16','#fff3e8','#e27e20'],
  ['rose','Rose','#9c365a','#fff0f4','#c24f78'],
  ['teal','Modern Teal','#087e83','#eaf8f7','#0b969b'],
  ['slate','Executive Slate','#34445a','#f1f4f8','#526780'],
  ['gold','Classic Gold','#806020','#fff9e9','#b58a36'],
  ['mono','Minimal Mono','#252525','#f7f7f7','#555555'],
  ['indigo','Indigo Night','#283b82','#eef0ff','#485fc7'],
  ['crimson','Crimson','#8f2636','#fff0f1','#bd3a4c'],
  ['bharat','Bharat','#9a4b13','#fff5e9','#ed8b2f']
 ];
 showDialog(`<h2>Document theme gallery</h2><p>Apply a coordinated heading colour, page tint and accent palette.</p><div class="dash-grid">${themes.map(t=>`<div class="template-card"><h3>${t[1]}</h3><div class="theme-swatch" style="background:linear-gradient(90deg,${t[2]} 0 33%,${t[3]} 33% 66%,${t[4]} 66%)"></div><p>Heading, paper and accent colours</p><button class="primary" onclick="applyCustomTheme('${t[0]}')">Apply theme</button></div>`).join('')}</div><div class="dialog-actions"><button onclick="closeDialog()">Close</button></div>`);
}
function applyCustomTheme(name){
 const map={ocean:['#174b91','#eaf2ff','#245bb5'],forest:['#176b45','#edf8f1','#16834c'],royal:['#583c91','#f3effb','#7652b8'],sunrise:['#a94f16','#fff3e8','#e27e20'],rose:['#9c365a','#fff0f4','#c24f78'],teal:['#087e83','#eaf8f7','#0b969b'],slate:['#34445a','#f1f4f8','#526780'],gold:['#806020','#fff9e9','#b58a36'],mono:['#252525','#f7f7f7','#555555'],indigo:['#283b82','#eef0ff','#485fc7'],crimson:['#8f2636','#fff0f1','#bd3a4c'],bharat:['#9a4b13','#fff5e9','#ed8b2f']};
 const c=map[name]||map.ocean;page.style.background=c[1];page.style.setProperty('--accent',c[2]);page.querySelectorAll('h1,h2,h3').forEach(h=>h.style.color=c[0]);document.documentElement.style.setProperty('--accent',c[2]);closeDialog();toast('Theme applied');
}
function insertCoverPage(){
 insertHTML(`<div style="min-height:760px;text-align:center;padding:110px 20px 40px;page-break-after:always"><div class="cover-kicker">[ORGANISATION / BRAND]</div><h1 class="cover-title">[DOCUMENT TITLE]</h1><p class="cover-subtitle">[Subtitle or description]</p><div style="height:150px"></div><div class="cover-meta">Prepared by: [Name]<br>[Department / Class]<br>[Date]</div></div><div style="page-break-before:always"></div>`);
}
function insertWatermark(){
 const t=prompt('Watermark text:','DRAFT');if(!t)return;
 page.querySelector('.watermark')?.remove();let w=document.createElement('div');w.className='watermark';w.textContent=t;page.prepend(w);toast('Watermark added');
}
function insertPageNumber(){insertHTML('<span class="page-number">[Page number]</span>');}
function styleTable(kind){
 let t=window.getSelection()?.anchorNode?.parentElement?.closest('table');
 if(!t){toast('Place the cursor inside a table first');return}
 t.classList.remove('table-blue','table-green','table-dark','table-minimal');t.classList.add('table-'+kind);toast('Table style applied');
}
function insertCallout(){insertHTML('<div class="callout"><b>Key point</b><p>Write an important note here.</p></div><p></p>')}
function insertChecklist(){insertHTML('<ul style="list-style-type:none"><li>☐ Task one</li><li>☐ Task two</li><li>☐ Task three</li></ul>')}
function setParagraphSpacing(){
 const v=prompt('Paragraph spacing after (px):','10');if(v===null)return;let n=Math.max(0,Math.min(48,parseInt(v)||0));page.querySelectorAll('p').forEach(p=>p.style.marginBottom=n+'px');toast('Paragraph spacing updated');
}
function setPageColumns(){
 const n=prompt('Number of columns (1–3):','2');if(n===null)return;let c=Math.max(1,Math.min(3,parseInt(n)||1));page.style.columnCount=c;page.style.columnGap='34px';toast(c+' column layout applied');
}
function clearPageColumns(){page.style.columnCount='';page.style.columnGap='';toast('Single-column layout restored')}
function insertSignatureLine(){insertHTML('<p style="margin-top:50px">____________________________<br>[Signature / Name]</p>')}


const TEXT_STYLE_PRESETS = [
 ['Classic Title','font-family:Georgia,serif;font-size:30pt;font-weight:700;color:#1f3556;letter-spacing:.3px'],
 ['Modern Title','font-family:Arial,sans-serif;font-size:28pt;font-weight:800;color:#174b91'],
 ['Elegant Title','font-family:Georgia,serif;font-size:27pt;font-style:italic;color:#704b76'],
 ['Bold Impact','font-family:Arial,sans-serif;font-size:25pt;font-weight:900;color:#202938;text-transform:uppercase'],
 ['Blue Heading','font-size:20pt;font-weight:700;color:#245bb5'],
 ['Green Heading','font-size:20pt;font-weight:700;color:#16834c'],
 ['Purple Heading','font-size:20pt;font-weight:700;color:#7652b8'],
 ['Red Heading','font-size:20pt;font-weight:700;color:#bd3a4c'],
 ['Gold Heading','font-size:20pt;font-weight:700;color:#a77a22'],
 ['Teal Heading','font-size:20pt;font-weight:700;color:#087e83'],
 ['Small Caps','font-variant:small-caps;font-size:16pt;font-weight:700;letter-spacing:1px;color:#34445a'],
 ['Section Label','font-size:10pt;font-weight:800;text-transform:uppercase;letter-spacing:2px;color:#526780'],
 ['Subtitle','font-family:Georgia,serif;font-size:16pt;color:#64748b;font-style:italic'],
 ['Lead Paragraph','font-size:14pt;line-height:1.7;color:#40536c'],
 ['Emphasis','font-weight:700;color:#245bb5'],
 ['Strong Highlight','font-weight:800;color:#9c365a;background:#fff0f4;padding:2px 5px'],
 ['Soft Highlight','background:#fff3b8;padding:2px 5px'],
 ['Code Inline','font-family:Consolas,monospace;background:#eef1f5;color:#9b3b20;padding:3px 6px;border-radius:4px'],
 ['Quote Accent','font-family:Georgia,serif;font-size:15pt;font-style:italic;color:#526780'],
 ['Number Accent','font-size:24pt;font-weight:800;color:#ed8b2f'],
 ['Caption','font-size:9pt;color:#778397;font-style:italic'],
 ['Muted Note','font-size:10pt;color:#68778d'],
 ['Warning','font-weight:700;color:#a33;background:#fff0ed;padding:3px 6px'],
 ['Success','font-weight:700;color:#16834c;background:#eaf7ef;padding:3px 6px'],
 ['Info','font-weight:700;color:#245bb5;background:#edf4ff;padding:3px 6px'],
 ['Kicker','font-size:9pt;font-weight:800;letter-spacing:3px;text-transform:uppercase;color:#e27e20'],
 ['Underline Accent','font-weight:700;text-decoration:underline;text-decoration-color:#ed8b2f;text-decoration-thickness:3px'],
 ['Double Underline','font-weight:700;text-decoration:underline double;color:#34445a'],
 ['Strike Note','color:#8a6470;text-decoration:line-through'],
 ['Superscript Look','font-size:10pt;vertical-align:super;color:#245bb5'],
 ['Monospace','font-family:Consolas,monospace;font-size:11pt;color:#334155'],
 ['Serif Body','font-family:Georgia,serif;font-size:12pt;color:#29364a'],
 ['Rounded Modern','font-family:Arial,sans-serif;font-size:14pt;font-weight:600;color:#087e83'],
 ['Royal','font-family:Georgia,serif;font-size:18pt;font-weight:700;color:#583c91'],
 ['Bharat Saffron','font-size:18pt;font-weight:800;color:#ed8b2f'],
 ['Bharat Green','font-size:18pt;font-weight:800;color:#16834c'],
 ['Bharat Navy','font-size:18pt;font-weight:800;color:#14243d'],
 ['Rose Elegant','font-family:Georgia,serif;font-size:17pt;color:#c24f78'],
 ['Forest','font-size:17pt;font-weight:700;color:#176b45'],
 ['Ocean','font-size:17pt;font-weight:700;color:#174b91'],
 ['Slate','font-size:17pt;font-weight:700;color:#34445a'],
 ['Gold','font-size:17pt;font-weight:700;color:#806020'],
 ['Centered Banner','display:inline-block;padding:6px 14px;background:#eaf2ff;color:#174b91;font-size:16pt;font-weight:800;border-radius:5px'],
 ['Pill Label','display:inline-block;padding:4px 10px;background:#e6f3ec;color:#185c3b;font-size:10pt;font-weight:800;border-radius:20px'],
 ['Outline Look','font-size:19pt;font-weight:800;color:#fff;text-shadow:-1px -1px 0 #245bb5,1px -1px 0 #245bb5,-1px 1px 0 #245bb5,1px 1px 0 #245bb5'],
 ['Shadow Title','font-size:22pt;font-weight:900;color:#245bb5;text-shadow:2px 2px 3px #aab8cc'],
 ['Gradient-like','font-size:20pt;font-weight:900;color:#7652b8;text-decoration:underline;text-decoration-color:#ed8b2f'],
 ['Letter Spaced','font-size:14pt;font-weight:700;letter-spacing:4px;color:#526780'],
 ['Compact Label','font-size:9pt;font-weight:800;letter-spacing:1px;text-transform:uppercase;color:#245bb5'],
 ['Editorial','font-family:Georgia,serif;font-size:21pt;font-weight:700;color:#26364d'],,
 ['Aurora','font-family:Georgia,serif;font-size:23pt;font-weight:700;color:#174b91;letter-spacing:.4px'],
 ['Midnight','font-family:Arial,sans-serif;font-size:18pt;font-weight:800;color:#16834c;text-transform:uppercase'],
 ['Cobalt','font-size:14pt;font-weight:700;color:#7652b8;letter-spacing:1px'],
 ['Lagoon','font-size:12pt;font-style:italic;color:#e27e20;border-left:3px solid #e27e20;padding-left:8px'],
 ['Mint','font-size:16pt;font-weight:800;color:#c24f78;text-shadow:1px 1px 2px #cbd5e1'],
 ['Coral','font-family:Georgia,serif;font-size:13pt;color:#087e83;line-height:1.7'],
 ['Lavender','font-size:10pt;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:#34445a'],
 ['Berry','display:inline-block;padding:5px 10px;border-radius:5px;background:#806020;color:#fff;font-weight:800'],
 ['Ruby','font-size:20pt;font-weight:900;color:#bd3a4c;text-decoration:underline;text-decoration-thickness:2px'],
 ['Amber','font-size:11pt;color:#526780;background:#52678018;padding:4px 7px;border-radius:4px'],
 ['Citrus','font-family:Georgia,serif;font-size:23pt;font-weight:700;color:#174b91;letter-spacing:.4px'],
 ['Graphite','font-family:Arial,sans-serif;font-size:18pt;font-weight:800;color:#16834c;text-transform:uppercase'],
 ['Cloud','font-size:14pt;font-weight:700;color:#7652b8;letter-spacing:1px'],
 ['Sapphire','font-size:12pt;font-style:italic;color:#e27e20;border-left:3px solid #e27e20;padding-left:8px'],
 ['Emerald','font-size:16pt;font-weight:800;color:#c24f78;text-shadow:1px 1px 2px #cbd5e1'],
 ['Amethyst','font-family:Georgia,serif;font-size:13pt;color:#087e83;line-height:1.7'],
 ['Copper','font-size:10pt;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:#34445a'],
 ['Pearl','display:inline-block;padding:5px 10px;border-radius:5px;background:#806020;color:#fff;font-weight:800'],
 ['Indigo','font-size:20pt;font-weight:900;color:#bd3a4c;text-decoration:underline;text-decoration-thickness:2px'],
 ['Turquoise','font-size:11pt;color:#526780;background:#52678018;padding:4px 7px;border-radius:4px'],
 ['Crimson','font-family:Georgia,serif;font-size:23pt;font-weight:700;color:#174b91;letter-spacing:.4px'],
 ['Sandstone','font-family:Arial,sans-serif;font-size:18pt;font-weight:800;color:#16834c;text-transform:uppercase'],
 ['Arctic','font-size:14pt;font-weight:700;color:#7652b8;letter-spacing:1px'],
 ['Tropical','font-size:12pt;font-style:italic;color:#e27e20;border-left:3px solid #e27e20;padding-left:8px'],
 ['Vintage','font-size:16pt;font-weight:800;color:#c24f78;text-shadow:1px 1px 2px #cbd5e1'],
 ['Minimal','font-family:Georgia,serif;font-size:13pt;color:#087e83;line-height:1.7'],
 ['Modern','font-size:10pt;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:#34445a'],
 ['Classic','display:inline-block;padding:5px 10px;border-radius:5px;background:#806020;color:#fff;font-weight:800'],
 ['Studio','font-size:20pt;font-weight:900;color:#bd3a4c;text-decoration:underline;text-decoration-thickness:2px'],
 ['Magazine','font-size:11pt;color:#526780;background:#52678018;padding:4px 7px;border-radius:4px'],
 ['Academic','font-family:Georgia,serif;font-size:23pt;font-weight:700;color:#174b91;letter-spacing:.4px'],
 ['Business','font-family:Arial,sans-serif;font-size:18pt;font-weight:800;color:#16834c;text-transform:uppercase'],
 ['Portfolio','font-size:14pt;font-weight:700;color:#7652b8;letter-spacing:1px'],
 ['Creative','font-size:12pt;font-style:italic;color:#e27e20;border-left:3px solid #e27e20;padding-left:8px'],
 ['Elegant','font-size:16pt;font-weight:800;color:#c24f78;text-shadow:1px 1px 2px #cbd5e1'],
 ['Formal','font-family:Georgia,serif;font-size:13pt;color:#087e83;line-height:1.7'],
 ['Friendly','font-size:10pt;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:#34445a'],
 ['Technical','display:inline-block;padding:5px 10px;border-radius:5px;background:#806020;color:#fff;font-weight:800'],
 ['Digital','font-size:20pt;font-weight:900;color:#bd3a4c;text-decoration:underline;text-decoration-thickness:2px'],
 ['Retro','font-size:11pt;color:#526780;background:#52678018;padding:4px 7px;border-radius:4px'],
 ['Neon','font-family:Georgia,serif;font-size:23pt;font-weight:700;color:#174b91;letter-spacing:.4px'],
 ['Pastel','font-family:Arial,sans-serif;font-size:18pt;font-weight:800;color:#16834c;text-transform:uppercase'],
 ['Monochrome','font-size:14pt;font-weight:700;color:#7652b8;letter-spacing:1px'],
 ['Bold Block','font-size:12pt;font-style:italic;color:#e27e20;border-left:3px solid #e27e20;padding-left:8px'],
 ['Soft Serif','font-size:16pt;font-weight:800;color:#c24f78;text-shadow:1px 1px 2px #cbd5e1'],
 ['Handwritten','font-family:Georgia,serif;font-size:13pt;color:#087e83;line-height:1.7'],
 ['Condensed','font-size:10pt;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:#34445a'],
 ['Wide Letter','display:inline-block;padding:5px 10px;border-radius:5px;background:#806020;color:#fff;font-weight:800'],
 ['Outline Blue','font-size:20pt;font-weight:900;color:#bd3a4c;text-decoration:underline;text-decoration-thickness:2px'],
 ['Outline Gold','font-size:11pt;color:#526780;background:#52678018;padding:4px 7px;border-radius:4px'],
 ['Shadow Navy','font-family:Georgia,serif;font-size:23pt;font-weight:700;color:#174b91;letter-spacing:.4px'],
 ['Shadow Rose','font-family:Arial,sans-serif;font-size:18pt;font-weight:800;color:#16834c;text-transform:uppercase'],
 ['Gradient Violet','font-size:14pt;font-weight:700;color:#7652b8;letter-spacing:1px'],
 ['Gradient Ocean','font-size:12pt;font-style:italic;color:#e27e20;border-left:3px solid #e27e20;padding-left:8px'],
 ['Gradient Sunset','font-size:16pt;font-weight:800;color:#c24f78;text-shadow:1px 1px 2px #cbd5e1'],
 ['Gradient Forest','font-family:Georgia,serif;font-size:13pt;color:#087e83;line-height:1.7'],
 ['Gradient Candy','font-size:10pt;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:#34445a'],
 ['Banner Blue','display:inline-block;padding:5px 10px;border-radius:5px;background:#806020;color:#fff;font-weight:800'],
 ['Banner Green','font-size:20pt;font-weight:900;color:#bd3a4c;text-decoration:underline;text-decoration-thickness:2px'],
 ['Banner Rose','font-size:11pt;color:#526780;background:#52678018;padding:4px 7px;border-radius:4px'],
 ['Banner Gold','font-family:Georgia,serif;font-size:23pt;font-weight:700;color:#174b91;letter-spacing:.4px'],
 ['Pill Blue','font-family:Arial,sans-serif;font-size:18pt;font-weight:800;color:#16834c;text-transform:uppercase'],
 ['Pill Purple','font-size:14pt;font-weight:700;color:#7652b8;letter-spacing:1px'],
 ['Pill Teal','font-size:12pt;font-style:italic;color:#e27e20;border-left:3px solid #e27e20;padding-left:8px'],
 ['Badge Orange','font-size:16pt;font-weight:800;color:#c24f78;text-shadow:1px 1px 2px #cbd5e1'],
 ['Badge Red','font-family:Georgia,serif;font-size:13pt;color:#087e83;line-height:1.7'],
 ['Badge Green','font-size:10pt;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:#34445a'],
 ['Quote Blue','display:inline-block;padding:5px 10px;border-radius:5px;background:#806020;color:#fff;font-weight:800'],
 ['Quote Gold','font-size:20pt;font-weight:900;color:#bd3a4c;text-decoration:underline;text-decoration-thickness:2px'],
 ['Quote Rose','font-size:11pt;color:#526780;background:#52678018;padding:4px 7px;border-radius:4px'],
 ['Quote Teal','font-family:Georgia,serif;font-size:23pt;font-weight:700;color:#174b91;letter-spacing:.4px'],
 ['Note Blue','font-family:Arial,sans-serif;font-size:18pt;font-weight:800;color:#16834c;text-transform:uppercase'],
 ['Note Yellow','font-size:14pt;font-weight:700;color:#7652b8;letter-spacing:1px'],
 ['Note Green','font-size:12pt;font-style:italic;color:#e27e20;border-left:3px solid #e27e20;padding-left:8px'],
 ['Note Pink','font-size:16pt;font-weight:800;color:#c24f78;text-shadow:1px 1px 2px #cbd5e1'],
 ['Alert Red','font-family:Georgia,serif;font-size:13pt;color:#087e83;line-height:1.7'],
 ['Alert Amber','font-size:10pt;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:#34445a'],
 ['Success Green','display:inline-block;padding:5px 10px;border-radius:5px;background:#806020;color:#fff;font-weight:800'],
 ['Info Cyan','font-size:20pt;font-weight:900;color:#bd3a4c;text-decoration:underline;text-decoration-thickness:2px'],
 ['Section Navy','font-size:11pt;color:#526780;background:#52678018;padding:4px 7px;border-radius:4px'],
 ['Section Teal','font-family:Georgia,serif;font-size:23pt;font-weight:700;color:#174b91;letter-spacing:.4px'],
 ['Section Plum','font-family:Arial,sans-serif;font-size:18pt;font-weight:800;color:#16834c;text-transform:uppercase'],
 ['Kicker Blue','font-size:14pt;font-weight:700;color:#7652b8;letter-spacing:1px'],
 ['Kicker Green','font-size:12pt;font-style:italic;color:#e27e20;border-left:3px solid #e27e20;padding-left:8px'],
 ['Kicker Rose','font-size:16pt;font-weight:800;color:#c24f78;text-shadow:1px 1px 2px #cbd5e1'],
 ['Number Blue','font-family:Georgia,serif;font-size:13pt;color:#087e83;line-height:1.7'],
 ['Number Green','font-size:10pt;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:#34445a'],
 ['Number Purple','display:inline-block;padding:5px 10px;border-radius:5px;background:#806020;color:#fff;font-weight:800'],
 ['Caption Blue','font-size:20pt;font-weight:900;color:#bd3a4c;text-decoration:underline;text-decoration-thickness:2px'],
 ['Caption Grey','font-size:11pt;color:#526780;background:#52678018;padding:4px 7px;border-radius:4px'],
 ['Caption Gold','font-family:Georgia,serif;font-size:23pt;font-weight:700;color:#174b91;letter-spacing:.4px'],
 ['Lead Serif','font-family:Arial,sans-serif;font-size:18pt;font-weight:800;color:#16834c;text-transform:uppercase'],
 ['Lead Sans','font-size:14pt;font-weight:700;color:#7652b8;letter-spacing:1px'],
 ['Body Elegant','font-size:12pt;font-style:italic;color:#e27e20;border-left:3px solid #e27e20;padding-left:8px'],
 ['Body Compact','font-size:16pt;font-weight:800;color:#c24f78;text-shadow:1px 1px 2px #cbd5e1'],
 ['Body Airy','font-family:Georgia,serif;font-size:13pt;color:#087e83;line-height:1.7'],
 ['Code Blue','font-size:10pt;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:#34445a'],
 ['Code Dark','display:inline-block;padding:5px 10px;border-radius:5px;background:#806020;color:#fff;font-weight:800'],
 ['Label Uppercase','font-size:20pt;font-weight:900;color:#bd3a4c;text-decoration:underline;text-decoration-thickness:2px'],
 ['Label Small Caps','font-size:11pt;color:#526780;background:#52678018;padding:4px 7px;border-radius:4px'],
 ['Editorial Serif','font-family:Georgia,serif;font-size:23pt;font-weight:700;color:#174b91;letter-spacing:.4px']
];
const EXTRA_STYLE_COLORS=['#174b91','#16834c','#7652b8','#c24f78','#087e83','#b26a16','#34445a','#9b3b20','#3c568a','#8b4b72'];
const EXTRA_STYLE_FONTS=['Georgia,serif','Arial,sans-serif','Verdana,sans-serif','Tahoma,sans-serif','Trebuchet MS,sans-serif','Courier New,monospace','Garamond,serif','Palatino Linotype,serif','Segoe UI,sans-serif','Impact,sans-serif'];
const EXTRA_STYLE_NAMES=['Aurora','Summit','Cobalt','Sage','Rose','Copper','Midnight','Ocean','Forest','Ruby'];
for(let i=0;i<100;i++){const c=EXTRA_STYLE_COLORS[i%EXTRA_STYLE_COLORS.length],f=EXTRA_STYLE_FONTS[(i*3)%EXTRA_STYLE_FONTS.length],n=EXTRA_STYLE_NAMES[i%EXTRA_STYLE_NAMES.length],size=12+(i%7)*2,bold=i%2?'700':'800',spacing=(i%4)*.4;TEXT_STYLE_PRESETS.push([`${n} Style ${i+1}`,`font-family:${f};font-size:${size}pt;font-weight:${bold};color:${c};letter-spacing:${spacing}px;${i%3===0?'font-style:italic;':''}${i%4===0?'text-transform:uppercase;':''}`]);}
const EXTRA_FONTS=['Alegreya','Alegreya Sans','Alegreya SC','Archivo','Archivo Black','Archivo Narrow','Asap','Barlow','Barlow Condensed','Barlow Semi Condensed','Bitter','Bree Serif','Cabin','Cairo','Cardo','Catamaran','Cinzel','Cormorant Garamond','Crimson Text','DM Sans','DM Serif Display','Dosis','EB Garamond','Exo 2','Fira Code','Fira Sans','Fjalla One','Fraunces','Heebo','IBM Plex Mono','IBM Plex Sans','IBM Plex Serif','Inconsolata','Inter','Josefin Sans','Karla','Lato Black','Libre Baskerville','Libre Franklin','Manrope','Merriweather','Merriweather Sans','Montserrat','Mulish','Nanum Gothic','Nunito','Nunito Sans','Old Standard TT','Open Sans','Oswald','Outfit','Overpass','Oxygen','PT Mono','PT Sans','PT Serif','Playfair Display','Plus Jakarta Sans','Poppins','Prompt','Public Sans','Quicksand','Rajdhani','Raleway','Red Hat Display','Roboto','Roboto Condensed','Roboto Mono','Roboto Slab','Rubik','Saira','Satisfy','Space Grotesk','Space Mono','Spectral','Tajawal','Titillium Web','Ubuntu','Urbanist','Vollkorn','Work Sans','Zilla Slab','Archivo Narrow','Arvo','Assistant','Barlow','Be Vietnam Pro','Bodoni Moda','Commissioner','Eczar','Gelasio','Hind','Kanit','Lexend','M PLUS 1','Mukta','Noto Sans Display','Noto Serif Display','Onest','PT Sans Narrow','Red Hat Text','Roboto Flex','Source Code Pro','Source Sans 3','Source Serif 4','Syne'];
FONT_FAMILIES.push(...EXTRA_FONTS);
function applyTextStyle(index){
 const item=TEXT_STYLE_PRESETS[index];if(!item)return;
 const [name,style]=item;let sel=window.getSelection();
 if(sel&&sel.rangeCount&&!sel.isCollapsed&&page.contains(sel.anchorNode)){
  const span=document.createElement('span');span.setAttribute('style',style);
  try{const range=sel.getRangeAt(0);range.surroundContents(span)}catch(e){document.execCommand('insertHTML',false,`<span style="${style}">${escapeHTML(sel.toString())}</span>`)}
 }else{
  document.execCommand('insertHTML',false,`<p><span style="${style}">${escapeHTML(name+' — type your text here')}</span></p>`);
 }
 updateStats();closeDialog();focusPage();toast(name+' style applied');
}
function showTextStyles(){
 const html=TEXT_STYLE_PRESETS.map((x,i)=>`<button class="gallery-item" onclick="applyTextStyle(${i})"><span class="gallery-preview" style="${x[1]}">Aa</span><small>${i+1}. ${x[0]}</small></button>`).join('');
 showDialog(`<h2>250+ Text Styles</h2><p>Select text first to style it, or choose a style to insert a styled sample paragraph.</p><div class="gallery-grid">${html}</div><div class="dialog-actions"><button onclick="closeDialog()">Close</button></div>`);
}
const BULLET_PRESETS=[
 ['•','Disc'],['○','Circle'],['▪','Square'],['◆','Diamond'],['➤','Arrow'],['➜','Long arrow'],['✓','Check'],['✔','Heavy check'],['☑','Checkbox'],['★','Star'],['☆','Outline star'],['✦','Sparkle'],['❖','Floral diamond'],['➢','Pointer'],['»','Double angle'],['→','Right arrow'],['♥','Heart'],['●','Large dot'],['–','Dash'],['☀','Sun']
];
function applyBulletStyle(index){
 const b=BULLET_PRESETS[index];if(!b)return;
 const lines=prompt('Enter list items, one item per line:','First item\\nSecond item\\nThird item');
 if(lines===null)return;
 const items=lines.split(/\\r?\\n/).filter(x=>x.trim()).map(x=>`<li style="list-style:none;margin:5px 0"><span style="display:inline-block;min-width:25px;color:#245bb5;font-weight:800">${b[0]}</span>${escapeHTML(x.trim())}</li>`).join('');
 insertHTML(`<ul style="padding-left:0">${items||`<li style="list-style:none">${b[0]} List item</li>`}</ul><p></p>`);
 closeDialog();toast(b[1]+' bullet list inserted');
}
function showBulletGallery(){
 const html=BULLET_PRESETS.map((b,i)=>`<button class="gallery-item" onclick="applyBulletStyle(${i})"><span class="gallery-preview">${b[0]}</span><small>${i+1}. ${b[1]}</small></button>`).join('');
 showDialog(`<h2>20 Bullet Styles</h2><p>Choose a marker, then enter list items on separate lines.</p><div class="gallery-grid">${html}</div><div class="dialog-actions"><button onclick="closeDialog()">Close</button></div>`);
}
const SHAPE_PRESETS=[
 ['Rectangle','▭'],['Rounded rectangle','▢'],['Square','□'],['Rounded square','▣'],['Circle','●'],['Oval','⬭'],['Ring','◎'],['Semicircle','◠'],['Triangle','▲'],['Right triangle','◢'],['Inverted triangle','▼'],['Diamond','◆'],['Parallelogram','▱'],['Trapezoid','⏢'],['Pentagon','⬟'],['Hexagon','⬡'],['Heptagon','⬢'],['Octagon','⯃'],['Decagon','⬠'],['Star 4 point','✦'],['Star 5 point','★'],['Star 6 point','✶'],['Star 8 point','✴'],['Star 12 point','✹'],['Burst','✺'],['Sun','☀'],['Moon','☾'],['Cloud','☁'],['Lightning','ϟ'],['Heart','♥'],
 ['Arrow right','➡'],['Arrow left','⬅'],['Arrow up','⬆'],['Arrow down','⬇'],['Left-right arrow','↔'],['Up-down arrow','↕'],['Double arrow','⇔'],['Chevron','❯'],['Pentagon arrow','➤'],['Circular arrow','↻'],['Curved arrow','⤴'],['U-turn arrow','↶'],['Notched arrow','➠'],['Callout arrow','➦'],['Line','／'],['Double line','═'],['Dashed line','┄'],['Wave','〰'],['Bracket pair','【】'],['Brace pair','{ }'],
 ['Flowchart process','▰'],['Flowchart decision','◇'],['Flowchart data','▱'],['Flowchart document','▤'],['Flowchart terminator','▱'],['Flowchart preparation','⬡'],['Flowchart connector','●'],['Flowchart storage','⌑'],['Flowchart delay','◖'],['Flowchart manual input','⏢'],
 ['Speech bubble','💬'],['Thought bubble','☁'],['Cloud callout','☁'],['Rectangular callout','▢'],['Rounded callout','▣'],['Brackets','［］'],['Parentheses','( )'],['Plus','＋'],['Minus','−'],['Multiply','×'],['Divide','÷'],['Equal','＝'],['Not equal','≠'],['Infinity','∞'],['Question mark','?'],['Exclamation','!'],['Check mark','✓'],['Cross mark','✕'],['Location pin','📍'],['Flag','⚑'],
  ['Star 16 point','✷'],['Star 24 point','✹'],['Outlined star','☆'],['Crescent','☽'],['Rainbow','🌈'],['Tree','♣'],['Flower','✿'],['Leaf','❧'],['Droplet','💧'],
  ['Gear','⚙'],['Crown','♛'],['Medal','🏅'],['Ribbon','🎗'],['Badge','✪'],['Shield','⬟'],['House','⌂'],['Envelope','✉'],['Folder','📁'],['Document','▤'],
  ['Book','▣'],['Calendar','▦'],['Clock','◷'],['Magnifier','⌕'],['Globe','◎'],['Network node','●'],['Database','▤'],['Cylinder','◉'],['Cube','⬡'],['Pyramid','▲'],
  ['Cone','△'],['Arc','⌒'],['Curve','∿'],['Loop','∞'],['Bracket left','〔'],['Bracket right','〕'],['Pentagon outline','⬠'],['Hexagon outline','⬡'],['Cross','✚'],['Four-way arrow','✥']
];
function shapeClip(name){const clips={Triangle:'polygon(50% 0,0 100%,100% 100%)','Inverted triangle':'polygon(0 0,100% 0,50% 100%)','Right triangle':'polygon(0 0,0 100%,100% 100%)',Diamond:'polygon(50% 0,100% 50%,50% 100%,0 50%)',Hexagon:'polygon(25% 0,75% 0,100% 50%,75% 100%,25% 100%,0 50%)',Pentagon:'polygon(50% 0,100% 38%,82% 100%,18% 100%,0 38%)',Parallelogram:'polygon(20% 0,100% 0,80% 100%,0 100%)',Trapezoid:'polygon(20% 0,80% 0,100% 100%,0 100%)',Octagon:'polygon(28% 0,72% 0,100% 28%,100% 72%,72% 100%,28% 100%,0 72%,0 28%)','Flowchart decision':'polygon(50% 0,100% 50%,50% 100%,0 50%)','Arrow right':'polygon(0 25%,65% 25%,65% 0,100% 50%,65% 100%,65% 75%,0 75%)','Arrow left':'polygon(100% 25%,35% 25%,35% 0,0 50%,35% 100%,35% 75%,100% 75%)','Arrow up':'polygon(25% 100%,25% 35%,0 35%,50% 0,100% 35%,75% 35%,75% 100%)','Arrow down':'polygon(25% 0,75% 0,75% 65%,100% 65%,50% 100%,0 65%,25% 65%)','Left-right arrow':'polygon(0 50%,22% 15%,22% 35%,78% 35%,78% 15%,100% 50%,78% 85%,78% 65%,22% 65%,22% 85%)','Up-down arrow':'polygon(50% 0,85% 22%,65% 22%,65% 78%,85% 78%,50% 100%,15% 78%,35% 78%,35% 22%,15% 22%)','Double arrow':'polygon(0 50%,20% 15%,20% 35%,80% 35%,80% 15%,100% 50%,80% 85%,80% 65%,20% 65%,20% 85%)','Chevron':'polygon(0 0,55% 0,100% 50%,55% 100%,0 100%,45% 50%)','Notched arrow':'polygon(0 0,70% 0,100% 50%,70% 100%,0 100%,25% 50%)','Star 5 point':'polygon(50% 0,61% 35%,98% 35%,68% 57%,79% 95%,50% 72%,21% 95%,32% 57%,2% 35%,39% 35%)','Star 4 point':'polygon(50% 0,62% 38%,100% 50%,62% 62%,50% 100%,38% 62%,0 50%,38% 38%)','Star 6 point':'polygon(50% 0,62% 30%,95% 25%,75% 50%,95% 75%,62% 70%,50% 100%,38% 70%,5% 75%,25% 50%,5% 25%,38% 30%)','Star 8 point':'polygon(50% 0,62% 32%,88% 12%,68% 38%,100% 50%,68% 62%,88% 88%,62% 68%,50% 100%,38% 68%,12% 88%,32% 62%,0 50%,32% 38%,12% 12%,38% 32%)','Heart':'polygon(50% 95%,8% 52%,5% 30%,18% 10%,38% 8%,50% 25%,62% 8%,82% 10%,95% 30%,92% 52%)','Cross':'polygon(35% 0,65% 0,65% 35%,100% 35%,100% 65%,65% 65%,65% 100%,35% 100%,35% 65%,0 65%,0 35%,35% 35%)','Plus':'polygon(40% 0,60% 0,60% 40%,100% 40%,100% 60%,60% 60%,60% 100%,40% 100%,40% 60%,0 60%,0 40%,40% 40%)','Flowchart data':'polygon(20% 0,100% 0,80% 100%,0 100%)','Speech bubble':'polygon(0 0,100% 0,100% 75%,65% 75%,45% 100%,45% 75%,0 75%)'};return clips[name]||''}
function insertShapePreset(index){
 const [name,glyph]=SHAPE_PRESETS[index]||SHAPE_PRESETS[0];const safe=escapeHTML(name);
 const colors=['#245bb5','#16834c','#7652b8','#e27e20','#c24f78','#087e83'];const color=colors[index%colors.length];
 const clip=shapeClip(name);const radius=['Circle','Oval','Ring'].includes(name)?'50%':name==='Rounded rectangle'?'18px':'5px';
 const ring=name==='Ring';const style=`width:170px;height:105px;background:${ring?'transparent':color};color:${ring?color:'#fff'};border:${ring?'8px':'3px'} solid ${color};border-radius:${radius};${clip?`clip-path:${clip};border:0;`:''}`;
 const text=name==='Line'||name==='Double line'||name==='Dashed line'?glyph:(name==='Rectangle'||name==='Rounded rectangle'||name==='Square'?'Edit text':safe);
 insertHTML(`<div contenteditable="true" spellcheck="true" title="${safe}" data-shape="${safe}" class="kms-shape" style="${style}">${escapeHTML(text)}</div><p></p>`);
 closeDialog();toast(safe+' inserted as a geometric shape — click to edit its label');
}
function showShapeGallery(){
 const html=SHAPE_PRESETS.map((s,i)=>`<button class="gallery-item" onclick="insertShapePreset(${i})"><span class="gallery-preview">${s[1]}</span><small>${i+1}. ${s[0]}</small></button>`).join('');
 showDialog(`<h2>100+ Geometric Shapes — editable labels, colours and sizes</h2><p>Choose from basic geometry, arrows, flowchart symbols, callouts and common symbols.</p><div class="gallery-grid">${html}</div><div class="dialog-actions"><button onclick="closeDialog()">Close</button></div>`);
}

let selectedObject=null, objectMoveMode=false;
function selectDocumentObject(el){
 if(selectedObject)selectedObject.classList.remove('object-selected');
 selectedObject=el;
 if(selectedObject)selectedObject.classList.add('object-selected');
 renderObjectInspector();
}
function renderObjectInspector(){
 const hint=document.getElementById('objectHint'),box=document.getElementById('objectControls');if(!box)return;
 if(!selectedObject){hint.textContent='Select a shape or picture on the page to edit it.';box.innerHTML='';return}
 const isImage=selectedObject.tagName==='IMG';hint.textContent=isImage?'Picture selected — edit size, crop, colour effects and border.':'Shape selected — edit its label, fill, outline, size and position.';
 const w=Math.round(selectedObject.getBoundingClientRect().width),h=Math.round(selectedObject.getBoundingClientRect().height),cs=getComputedStyle(selectedObject);
 if(isImage){box.innerHTML=`<label>Width (px)<input id="objWidth" type="number" min="20" max="1600" value="${w}"></label><label>Height (px)<input id="objHeight" type="number" min="20" max="2200" value="${h}"></label><label>Brightness <input id="imgBright" type="range" min="20" max="200" value="${selectedObject.dataset.bright||100}" oninput="applyImageFilters()"></label><label>Contrast <input id="imgContrast" type="range" min="20" max="200" value="${selectedObject.dataset.contrast||100}" oninput="applyImageFilters()"></label><label>Saturation <input id="imgSaturate" type="range" min="0" max="200" value="${selectedObject.dataset.saturate||100}" oninput="applyImageFilters()"></label><label>Crop / fit <select id="imgFit" onchange="applyImageFit()"><option value="contain">Fit inside</option><option value="cover">Fill / crop</option></select></label><label>Crop horizontal <input id="imgCropX" type="range" min="0" max="100" value="50" oninput="applyImageCrop()"></label><label>Crop vertical <input id="imgCropY" type="range" min="0" max="100" value="50" oninput="applyImageCrop()"></label><label>Border colour<input id="objBorder" type="color" value="${rgbToHex(cs.borderTopColor)}"></label><label>Border width (px)<input id="imgBorderWidth" type="number" min="0" max="30" value="${parseInt(cs.borderTopWidth)||0}"></label><div><button onclick="applyObjectSize()">Apply size</button><button onclick="applyImageBorder()">Apply border</button></div><div><button onclick="flipImage('x')">Flip horizontal</button><button onclick="flipImage('y')">Flip vertical</button></div><div><button onclick="rotateObject()">Rotate 15°</button><button onclick="deleteSelectedObject()" style="color:#a33">Delete</button></div>`;
 }else{box.innerHTML=`<label>Width (px)<input id="objWidth" type="number" min="20" max="1600" value="${w}"></label><label>Height (px)<input id="objHeight" type="number" min="20" max="2200" value="${h}"></label><label>Fill<input id="objFill" type="color" value="${rgbToHex(cs.backgroundColor)}"></label><label>Text colour<input id="objText" type="color" value="${rgbToHex(cs.color)}"></label><label>Border colour<input id="objBorder" type="color" value="${rgbToHex(cs.borderTopColor)}"></label><label>Border width (px)<input id="objBorderWidth" type="number" min="0" max="30" value="${parseInt(cs.borderTopWidth)||0}"></label><label>Corner radius (px)<input id="objRadius" type="number" min="0" max="100" value="${parseInt(cs.borderRadius)||5}"></label><div><button onclick="applyObjectSize()">Apply size</button><button onclick="applyObjectColours()">Apply colours</button></div><div><button onclick="toggleObjectMove()">${objectMoveMode?'Stop moving':'Move object'}</button><button onclick="rotateObject()">Rotate 15°</button><button onclick="deleteSelectedObject()" style="color:#a33">Delete</button></div>`;}
}
function rgbToHex(v){const m=String(v).match(/\d+/g);if(!m||m.length<3)return '#245bb5';return '#'+m.slice(0,3).map(x=>Number(x).toString(16).padStart(2,'0')).join('')}
function applyObjectSize(){if(!selectedObject)return;let w=Math.max(20,Math.min(1600,Number(document.getElementById('objWidth').value)||150)),h=Math.max(20,Math.min(2200,Number(document.getElementById('objHeight').value)||80));selectedObject.style.width=w+'px';selectedObject.style.height=h+'px';if(selectedObject.tagName==='IMG')selectedObject.style.objectFit='contain';renderObjectInspector()}
function applyObjectColours(){if(!selectedObject||selectedObject.tagName==='IMG')return;selectedObject.style.backgroundColor=document.getElementById('objFill').value;selectedObject.style.color=document.getElementById('objText').value;selectedObject.style.borderColor=document.getElementById('objBorder').value;selectedObject.style.setProperty('--shape-border',document.getElementById('objBorder').value);selectedObject.style.borderWidth=(Number(document.getElementById('objBorderWidth').value)||2)+'px';selectedObject.style.borderRadius=(Number(document.getElementById('objRadius').value)||0)+'px'}
function applyImageFilters(){if(!selectedObject||selectedObject.tagName!=='IMG')return;let b=document.getElementById('imgBright').value,c=document.getElementById('imgContrast').value,s=document.getElementById('imgSaturate').value;selectedObject.dataset.bright=b;selectedObject.dataset.contrast=c;selectedObject.dataset.saturate=s;selectedObject.style.filter=`brightness(${b}%) contrast(${c}%) saturate(${s}%)`}
function applyImageFit(){if(selectedObject&&selectedObject.tagName==='IMG')selectedObject.style.objectFit=document.getElementById('imgFit').value}
function applyImageCrop(){if(selectedObject&&selectedObject.tagName==='IMG')selectedObject.style.objectPosition=document.getElementById('imgCropX').value+'% '+document.getElementById('imgCropY').value+'%'}
function applyImageBorder(){if(!selectedObject||selectedObject.tagName!=='IMG')return;selectedObject.style.border=`${Math.max(0,Number(document.getElementById('imgBorderWidth').value)||0)}px solid ${document.getElementById('objBorder').value}`;selectedObject.style.borderRadius='4px'}
function flipImage(axis){if(!selectedObject||selectedObject.tagName!=='IMG')return;selectedObject.dataset['flip'+axis]=(selectedObject.dataset['flip'+axis]==='1'?'0':'1');const x=selectedObject.dataset.flipx==='1'?-1:1,y=selectedObject.dataset.flipy==='1'?-1:1;selectedObject.style.transform=`scale(${x},${y}) rotate(${Number(selectedObject.dataset.rotation||0)}deg)`}
function rotateObject(){if(!selectedObject)return;let a=Number(selectedObject.dataset.rotation||0)+15;selectedObject.dataset.rotation=a;let flipX=selectedObject.dataset.flipx==='1'?-1:1,flipY=selectedObject.dataset.flipy==='1'?-1:1;selectedObject.style.transform=`scale(${flipX},${flipY}) rotate(${a}deg)`}
function deleteSelectedObject(){if(!selectedObject)return;selectedObject.remove();selectedObject=null;renderObjectInspector();updateStats()}
function toggleObjectMove(){objectMoveMode=!objectMoveMode;renderObjectInspector();toast(objectMoveMode?'Move mode: drag the selected object':'Move mode off')}
page.addEventListener('click',e=>{const obj=e.target.closest('.kms-shape,.kms-image,img');if(obj&&page.contains(obj)){if(obj.tagName==='IMG')obj.classList.add('kms-image');selectDocumentObject(obj)}else if(!e.target.closest('#objectInspector'))selectDocumentObject(null)});
page.addEventListener('pointerdown',e=>{if(!objectMoveMode||!selectedObject||!selectedObject.contains(e.target))return;e.preventDefault();const el=selectedObject,startX=e.clientX,startY=e.clientY,baseX=Number(el.dataset.moveX||0),baseY=Number(el.dataset.moveY||0);const move=ev=>{const dx=baseX+ev.clientX-startX,dy=baseY+ev.clientY-startY;el.dataset.moveX=dx;el.dataset.moveY=dy;const fx=el.dataset.flipx==='1'?-1:1,fy=el.dataset.flipy==='1'?-1:1;el.style.transform=`translate(${dx}px,${dy}px) scale(${fx},${fy}) rotate(${Number(el.dataset.rotation||0)}deg)`};const up=()=>{document.removeEventListener('pointermove',move);document.removeEventListener('pointerup',up)};document.addEventListener('pointermove',move);document.addEventListener('pointerup',up,{once:true})});

function openWritingAssistant(){showDialog(`<h2>✦ KMS Writing Assistant</h2><p>This offline assistant offers quick editing actions. It does not send your document to an online service.</p><label>Text to work on<textarea id="aiInput" rows="7" placeholder="Paste or type text here…"></textarea></label><div class="gallery-grid"><button class="gallery-item" onclick="runLocalAI('formal')">Make formal</button><button class="gallery-item" onclick="runLocalAI('clear')">Improve clarity</button><button class="gallery-item" onclick="runLocalAI('summary')">Short summary</button><button class="gallery-item" onclick="runLocalAI('grammar')">Punctuation check</button><button class="gallery-item" onclick="runLocalAI('bullets')">Convert to bullets</button><button class="gallery-item" onclick="runLocalAI('outline')">Create outline</button></div><label>Result<textarea id="aiOutput" rows="7" placeholder="Your result appears here…"></textarea></label><div class="dialog-actions"><button onclick="insertAIResult()">Insert into document</button><button class="primary" onclick="closeDialog()">Close</button></div>`)}
function insertAIResult(){const v=document.getElementById('aiOutput').value;if(!v){toast('Create a result first');return}closeDialog();focusPage();insertHTML('<p>'+escapeHTML(v).replace(/\n/g,'<br>')+'</p>')}
function runLocalAI(mode){
 let v=document.getElementById('aiInput').value.trim(),o=v;
 if(!v){toast('Enter some text first');return}
 if(mode==='formal') o=v.replace(/\bget\b/gi,'obtain').replace(/\ba lot of\b/gi,'many').replace(/\bshow\b/gi,'demonstrate').replace(/\bhelp\b/gi,'assist').replace(/\bneed to\b/gi,'should').replace(/\bcan't\b/gi,'cannot');
 else if(mode==='clear') o=v.replace(/\bin order to\b/gi,'to').replace(/\bdue to the fact that\b/gi,'because').replace(/\bat this point in time\b/gi,'now').replace(/\butilize\b/gi,'use');
 else if(mode==='summary'){let sentences=v.match(/[^.!?]+[.!?]*/g)||[v];o=sentences.slice(0,Math.max(1,Math.ceil(sentences.length/3))).join(' ').trim()}
 else if(mode==='grammar'){o=v.replace(/\s+([,.!?;:])/g,'$1').replace(/([,.!?])(?=[A-Za-z])/g,'$1 ').replace(/\bi\b/g,'I').replace(/\s{2,}/g,' ');if(o&&!/[.!?]$/.test(o))o+='.'}
 else if(mode==='bullets') o=v.split(/[\n.!?]+/).map(x=>x.trim()).filter(Boolean).map(x=>'• '+x).join('\n');
 else if(mode==='outline') o='Overview\n'+v.split(/[.!?]+/).map(x=>x.trim()).filter(Boolean).map((x,i)=>(i+1)+'. '+x).join('\n');
 document.getElementById('aiOutput').value=o;
}
function showDialog(html){document.getElementById('dialog').innerHTML=html;document.getElementById('overlay').style.display='flex'}
function closeDialog(){document.getElementById('overlay').style.display='none'}
function showGuide(){showDialog('<h2>KMS WORD AI V20.0 PRO — Quick guide</h2><p><b>1.</b> Type directly on the white page.</p><p><b>2.</b> Use Home for font and paragraph formatting.</p><p><b>3.</b> Use Insert to add tables, pictures, links, shapes and page breaks.</p><p><b>4.</b> Layout and Design control page appearance.</p><p><b>5.</b> Documents autosave in this browser profile. Use Save to save immediately.</p><p><b>6.</b> Export HTML for an editable web document, TXT for plain text, or Print → Save as PDF.</p><p><b>Important:</b> Saved documents are stored locally on this computer/browser profile. Export backups regularly.</p><div class="dialog-actions"><button class="primary" onclick="closeDialog()">Close</button></div>')}
function showShortcuts(){showDialog('<h2>Keyboard shortcuts</h2><p><b>Ctrl+B</b> Bold · <b>Ctrl+I</b> Italic · <b>Ctrl+U</b> Underline</p><p><b>Ctrl+Z</b> Undo · <b>Ctrl+Y</b> Redo · <b>Ctrl+S</b> Save</p><p><b>Ctrl+F</b> Find · <b>Ctrl+P</b> Print / PDF · <b>Ctrl+A</b> Select all</p><div class="dialog-actions"><button class="primary" onclick="closeDialog()">Close</button></div>')}
function aboutApp(){showDialog('<h2>KMS WORD AI — Version 20.0 PRO</h2><p>Advanced Office-style document editor by KMS A VISION. Version 20.0 PRO adds improved startup resilience, themes, page design and productivity tools.</p><p>Desktop shell: Python + PyQt6 + Qt WebEngine.</p><p>Document content and saved documents are stored locally in the application web profile.</p><div class="dialog-actions"><button class="primary" onclick="closeDialog()">Close</button></div>')}
let boardCtx=null,boardHistory=[],boardRedoStack=[],boardStart=null,boardZoomLevel=1,boardGrid=false,boardSnapshotTimer=null;
function openWhiteboard(){document.getElementById('whiteboard').style.display='flex';setTimeout(()=>{resizeBoard();if(!boardHistory.length)saveBoardHistory()},50)}
function closeWhiteboard(){document.getElementById('whiteboard').style.display='none'}
const canvas=document.getElementById('board'),ctx=canvas.getContext('2d');boardCtx=ctx;
function resizeBoard(){const stage=document.querySelector('.boardstage'),r=stage.getBoundingClientRect(),d=window.devicePixelRatio||1;if(!r.width||!r.height)return;const old=document.createElement('canvas');old.width=canvas.width;old.height=canvas.height;old.getContext('2d').drawImage(canvas,0,0);canvas.width=r.width*d;canvas.height=r.height*d;ctx.setTransform(d,0,0,d,0,0);if(old.width&&old.height)ctx.drawImage(old,0,0,r.width,r.height);if(boardGrid)drawBoardGrid()}
function boardPos(e){let r=canvas.getBoundingClientRect(),x=(e.clientX-r.left)/boardZoomLevel,y=(e.clientY-r.top)/boardZoomLevel;if(boardSnap){x=Math.round(x/10)*10;y=Math.round(y/10)*10}return [x,y]}
function setBoardTool(t){boardTool=t;document.getElementById('boardToolStatus').textContent='Tool: '+t;canvas.style.cursor=t==='text'?'text':t==='eraser'?'cell':'crosshair'}
function saveBoardHistory(){try{boardHistory.push(canvas.toDataURL('image/png'));if(boardHistory.length>35)boardHistory.shift();boardRedoStack=[]}catch(e){}}
function restoreBoardData(url){const im=new Image();im.onload=()=>{ctx.clearRect(0,0,canvas.width,canvas.height);ctx.drawImage(im,0,0,canvas.clientWidth,canvas.clientHeight)};im.src=url}
function boardUndo(){if(boardHistory.length<2)return;boardRedoStack.push(boardHistory.pop());restoreBoardData(boardHistory[boardHistory.length-1])}
function boardRedo(){if(!boardRedoStack.length)return;const u=boardRedoStack.pop();boardHistory.push(u);restoreBoardData(u)}
function drawBoardGrid(){if(!boardGrid)return;const w=canvas.clientWidth,h=canvas.clientHeight;ctx.save();ctx.strokeStyle='#dce4ef';ctx.lineWidth=.6;for(let x=20;x<w;x+=20){ctx.beginPath();ctx.moveTo(x,0);ctx.lineTo(x,h);ctx.stroke()}for(let y=20;y<h;y+=20){ctx.beginPath();ctx.moveTo(0,y);ctx.lineTo(w,y);ctx.stroke()}ctx.restore()}
function toggleBoardGrid(){boardGrid=!boardGrid;toast(boardGrid?'Grid on':'Grid off');if(boardGrid){drawBoardGrid()}else{restoreBoardData(boardHistory[boardHistory.length-1]||canvas.toDataURL())}}
function setBoardBackground(c){ctx.save();ctx.globalCompositeOperation='destination-over';ctx.fillStyle=c;ctx.fillRect(0,0,canvas.width,canvas.height);ctx.restore();saveBoardHistory()}
function boardZoom(n){boardZoomLevel=Math.max(.5,Math.min(2,boardZoomLevel+n/100));document.getElementById('boardZoomStatus').textContent='Zoom: '+Math.round(boardZoomLevel*100)+'%';canvas.style.transform=`scale(${boardZoomLevel})`;canvas.style.transformOrigin='top left'}
function boardFit(){boardZoomLevel=1;canvas.style.transform='none';document.getElementById('boardZoomStatus').textContent='Zoom: 100%'}
function boardShape(x1,y1,x2,y2,tool){ctx.beginPath();if(tool==='line'||tool==='arrow'){ctx.moveTo(x1,y1);ctx.lineTo(x2,y2);ctx.stroke();if(tool==='arrow'){let a=Math.atan2(y2-y1,x2-x1);ctx.beginPath();ctx.moveTo(x2,y2);ctx.lineTo(x2-14*Math.cos(a-.45),y2-14*Math.sin(a-.45));ctx.moveTo(x2,y2);ctx.lineTo(x2-14*Math.cos(a+.45),y2-14*Math.sin(a+.45));ctx.stroke()}}else if(tool==='rect')ctx.strokeRect(x1,y1,x2-x1,y2-y1);else if(tool==='ellipse'){ctx.ellipse((x1+x2)/2,(y1+y2)/2,Math.abs(x2-x1)/2,Math.abs(y2-y1)/2,0,0,Math.PI*2);ctx.stroke()}else if(tool==='triangle'){ctx.moveTo((x1+x2)/2,y1);ctx.lineTo(x2,y2);ctx.lineTo(x1,y2);ctx.closePath();ctx.stroke()}else if(tool==='diamond'){ctx.moveTo((x1+x2)/2,y1);ctx.lineTo(x2,(y1+y2)/2);ctx.lineTo((x1+x2)/2,y2);ctx.lineTo(x1,(y1+y2)/2);ctx.closePath();ctx.stroke()}else if(tool==='star'){let cx=(x1+x2)/2,cy=(y1+y2)/2,rx=Math.abs(x2-x1)/2,ry=Math.abs(y2-y1)/2;for(let i=0;i<10;i++){let a=-Math.PI/2+i*Math.PI/5,rad=i%2?0.42:1,xx=cx+Math.cos(a)*rx*rad,yy=cy+Math.sin(a)*ry*rad;i?ctx.lineTo(xx,yy):ctx.moveTo(xx,yy)}ctx.closePath();ctx.stroke()}}
canvas.addEventListener('pointerdown',e=>{if(e.button!==0)return;const [x,y]=boardPos(e);boardStart=[x,y];if(boardTool==='text'){let t=prompt('Text to add to board:','Your text');if(t){ctx.fillStyle=document.getElementById('boardColor').value;ctx.font='20px Segoe UI';ctx.fillText(t,x,y);saveBoardHistory()}return}if(boardTool==='sticky'){addStickyNote();return}if(['pen','highlighter','eraser'].includes(boardTool)){drawing=true;lastX=x;lastY=y;canvas.setPointerCapture(e.pointerId)}else{drawing=true;canvas.setPointerCapture(e.pointerId);boardSnapshotTimer=canvas.toDataURL('image/png')}});
canvas.addEventListener('pointermove',e=>{if(!drawing)return;const [x,y]=boardPos(e),color=document.getElementById('boardColor').value,width=Number(document.getElementById('boardWidth').value);if(['pen','highlighter','eraser'].includes(boardTool)){ctx.beginPath();ctx.moveTo(lastX,lastY);ctx.lineTo(x,y);ctx.lineWidth=boardTool==='highlighter'?Math.max(width,14):width;ctx.lineCap='round';ctx.globalAlpha=boardTool==='highlighter'?.25:1;ctx.globalCompositeOperation=boardTool==='eraser'?'destination-out':'source-over';ctx.strokeStyle=color;ctx.stroke();ctx.globalAlpha=1;ctx.globalCompositeOperation='source-over';lastX=x;lastY=y}});
function finishBoardDraw(e){if(!drawing)return;const [x,y]=boardPos(e);if(boardStart&&['line','arrow','rect','ellipse','triangle','diamond','star'].includes(boardTool)){if(boardSnapshotTimer){const im=new Image();im.onload=()=>{ctx.clearRect(0,0,canvas.width,canvas.height);ctx.drawImage(im,0,0,canvas.clientWidth,canvas.clientHeight);ctx.strokeStyle=document.getElementById('boardColor').value;ctx.lineWidth=Number(document.getElementById('boardWidth').value);boardShape(boardStart[0],boardStart[1],x,y,boardTool);saveBoardHistory()};im.src=boardSnapshotTimer}else{ctx.strokeStyle=document.getElementById('boardColor').value;ctx.lineWidth=Number(document.getElementById('boardWidth').value);boardShape(boardStart[0],boardStart[1],x,y,boardTool);saveBoardHistory()}}else saveBoardHistory();drawing=false;boardStart=null;boardSnapshotTimer=null}
canvas.addEventListener('pointerup',finishBoardDraw);canvas.addEventListener('pointercancel',finishBoardDraw);
function addStickyNote(){let t=prompt('Sticky note text:','Remember this idea');if(!t)return;let x=70+Math.random()*180,y=70+Math.random()*130;ctx.fillStyle='#fff1a8';ctx.fillRect(x,y,220,130);ctx.strokeStyle='#d7b84d';ctx.strokeRect(x,y,220,130);ctx.fillStyle='#4b3c16';ctx.font='16px Segoe UI';t.split('\n').slice(0,6).forEach((line,i)=>ctx.fillText(line.slice(0,25),x+12,y+28+i*22));saveBoardHistory()}
function chooseBoardImage(){document.getElementById('boardImageFile').click()}
function insertBoardImage(e){let f=e.target.files&&e.target.files[0];if(!f)return;let rd=new FileReader();rd.onload=()=>{let im=new Image();im.onload=()=>{let w=Math.min(im.width,500),h=im.height*w/im.width;ctx.drawImage(im,40,40,w,h);saveBoardHistory()};im.src=rd.result};rd.readAsDataURL(f);e.target.value=''}
function clearBoard(){if(!confirm('Clear the entire whiteboard?'))return;ctx.clearRect(0,0,canvas.width,canvas.height);saveBoardHistory()}
function downloadBoard(){let a=document.createElement('a');a.href=canvas.toDataURL('image/png');a.download='KMS_Whiteboard.png';a.click()}
function downloadBoardSVG(){const data=canvas.toDataURL('image/png');const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${canvas.width}" height="${canvas.height}"><image href="${data}" width="100%" height="100%"/></svg>`;let a=document.createElement('a');a.href=URL.createObjectURL(new Blob([svg],{type:'image/svg+xml'}));a.download='KMS_Whiteboard.svg';a.click()}
let boardSnap=false,boardDots=false,boardRuler=false,boardGuides=false,boardCrosshair=true;
function setBoardColor(c){document.getElementById('boardColor').value=c}
function setBoardWidth(w){document.getElementById('boardWidth').value=w}
function toggleBoardSnap(){boardSnap=!boardSnap;toast(boardSnap?'Snap to grid on':'Snap to grid off')}
function toggleBoardDots(){boardDots=!boardDots;if(boardDots){ctx.save();ctx.fillStyle='#d5deea';for(let x=10;x<canvas.clientWidth;x+=20)for(let y=10;y<canvas.clientHeight;y+=20){ctx.beginPath();ctx.arc(x,y,1,0,Math.PI*2);ctx.fill()}ctx.restore()}else restoreBoardData(boardHistory[boardHistory.length-1]||canvas.toDataURL())}
function toggleBoardRuler(){boardRuler=!boardRuler;if(boardRuler){ctx.save();ctx.strokeStyle='#8aa0bd';ctx.fillStyle='#4c6382';ctx.font='10px Arial';for(let x=0;x<canvas.clientWidth;x+=50){ctx.beginPath();ctx.moveTo(x,0);ctx.lineTo(x,14);ctx.stroke();ctx.fillText(String(x),x+2,25)}ctx.restore()}else restoreBoardData(boardHistory[boardHistory.length-1]||canvas.toDataURL())}
function toggleBoardGuides(){boardGuides=!boardGuides;ctx.save();ctx.strokeStyle='#8bb7ee';ctx.setLineDash([5,5]);if(boardGuides){ctx.beginPath();ctx.moveTo(canvas.clientWidth/2,0);ctx.lineTo(canvas.clientWidth/2,canvas.clientHeight);ctx.moveTo(0,canvas.clientHeight/2);ctx.lineTo(canvas.clientWidth,canvas.clientHeight/2);ctx.stroke()}else restoreBoardData(boardHistory[boardHistory.length-1]||canvas.toDataURL());ctx.restore()}
function toggleBoardCrosshair(){boardCrosshair=!boardCrosshair;canvas.style.cursor=boardCrosshair?'crosshair':'default'}
function downloadBoardJPG(){let a=document.createElement('a');a.href=canvas.toDataURL('image/jpeg',.92);a.download='KMS_Whiteboard.jpg';a.click()}
function downloadBoardJSON(){let data={version:1,width:canvas.width,height:canvas.height,image:canvas.toDataURL('image/png')};let a=document.createElement('a');a.href=URL.createObjectURL(new Blob([JSON.stringify(data)],{type:'application/json'}));a.download='KMS_Whiteboard.json';a.click()}
function saveBoardProject(){downloadBoardJSON()}
function loadBoardProject(){let i=document.createElement('input');i.type='file';i.accept='.json,application/json';i.onchange=()=>{let f=i.files[0];if(!f)return;let r=new FileReader();r.onload=()=>{try{let d=JSON.parse(r.result);restoreBoardData(d.image);saveBoardHistory()}catch(e){toast('Invalid whiteboard project file')}};r.readAsText(f)};i.click()}
function copyBoard(){canvas.toBlob(b=>{if(!b)return;navigator.clipboard?.write([new ClipboardItem({'image/png':b})]).then(()=>toast('Board copied')).catch(()=>toast('Clipboard permission unavailable'))})}
function saveBoardLocal(){try{localStorage.setItem('kms_whiteboard_local_v19',canvas.toDataURL('image/png'));toast('Whiteboard saved on this device')}catch(e){toast('Local storage unavailable — use PNG or JSON export')}}
function restoreBoardLocal(){try{let u=localStorage.getItem('kms_whiteboard_local_v19');if(!u){toast('No local board backup found');return}restoreBoardData(u);saveBoardHistory();toast('Local board restored')}catch(e){toast('Local storage unavailable — open a saved JSON project instead')}}
function mirrorBoard(){let im=new Image();im.onload=()=>{ctx.clearRect(0,0,canvas.width,canvas.height);ctx.save();ctx.translate(canvas.width,0);ctx.scale(-1,1);ctx.drawImage(im,0,0,canvas.width,canvas.height);ctx.restore();saveBoardHistory()};im.src=canvas.toDataURL()}
function rotateBoard90(){let im=new Image();im.onload=()=>{let tmp=document.createElement('canvas');tmp.width=canvas.height;tmp.height=canvas.width;let t=tmp.getContext('2d');t.translate(tmp.width/2,tmp.height/2);t.rotate(Math.PI/2);t.drawImage(im,-canvas.width/2,-canvas.height/2);ctx.clearRect(0,0,canvas.width,canvas.height);ctx.drawImage(tmp,0,0,canvas.width,canvas.height);saveBoardHistory()};im.src=canvas.toDataURL()}
function invertBoard(){let d=ctx.getImageData(0,0,canvas.width,canvas.height),a=d.data;for(let i=0;i<a.length;i+=4){a[i]=255-a[i];a[i+1]=255-a[i+1];a[i+2]=255-a[i+2]}ctx.putImageData(d,0,0);saveBoardHistory()}
function addBoardFrame(){ctx.save();ctx.strokeStyle=document.getElementById('boardColor').value;ctx.lineWidth=6;ctx.strokeRect(12,12,canvas.clientWidth-24,canvas.clientHeight-24);ctx.restore();saveBoardHistory()}
function newBoard(){ctx.clearRect(0,0,canvas.width,canvas.height);saveBoardHistory();toast('New whiteboard ready')}
let pendingEntryRange=null;
page.addEventListener('beforeinput',e=>{if(e.inputType&&e.inputType.startsWith('insert')){let sel=getSelection();if(sel&&sel.rangeCount)pendingEntryRange=sel.getRangeAt(0).cloneRange()}else pendingEntryRange=null});
function markNewEntry(){if(!pendingEntryRange)return;let sel=getSelection();if(!sel||!sel.rangeCount){pendingEntryRange=null;return}try{let end=sel.getRangeAt(0).cloneRange(),r=document.createRange();r.setStart(pendingEntryRange.startContainer,pendingEntryRange.startOffset);r.setEnd(end.endContainer,end.endOffset);if(r.collapsed||!page.contains(r.commonAncestorContainer)){pendingEntryRange=null;return}let span=document.createElement('span');span.className='temp-entry-highlight';r.surroundContents(span);setTimeout(()=>{if(!span.isConnected)return;span.style.backgroundColor='transparent';setTimeout(()=>{if(span.isConnected){let parent=span.parentNode;while(span.firstChild)parent.insertBefore(span.firstChild,span);parent.removeChild(span);}},5200)},55000)}catch(err){}pendingEntryRange=null}
page.addEventListener('input',()=>{markNewEntry();updateStats()});
document.getElementById('docTitle').addEventListener('input',scheduleAutosave);
document.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='s'){e.preventDefault();saveDocument()}if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='f'){e.preventDefault();openFind()}});
(function init(){
 // Build the ribbon first: storage restrictions must never prevent the tabs from working.
 try { setTab('home'); } catch(tabError) { console.error('Home ribbon failed:', tabError); document.getElementById('ribbon').textContent='Ribbon startup issue — select another tab or restart the app.'; }
 try {
  let id=null; try { id=localStorage.getItem(ACTIVE_KEY); } catch(storageError) { console.warn('Local storage unavailable; starting without restore.', storageError); }
  let d=id&&docs()[id];
  if(d){page.innerHTML=d.html;document.getElementById('docTitle').value=d.title||'Untitled Document'}
 } catch(initError) { console.error('Document restore failed:', initError); }
 try { updateStats(); } catch(statsError) { console.error('Stats initialization failed:', statsError); }
 try { renderRecent(); } catch(recentError) { console.error('Recent list initialization failed:', recentError); }
})();

// ================= KMS WORD AI V20 PRO EXTENSION =================
// 10 additional tabs, 300 ribbon commands and expanded sample galleries.
// AI editing below is offline/local utility logic, not a connected cloud model.
const PRO_FEATURES = {
 filetools: "New|Save|Export HTML|Export TXT|Print PDF|Duplicate|Rename|Properties|Open|Backup|Autosave|Cover page|Page break|Header|Footer|Page number|TOC|Bibliography|Footnote|Citation|Find|Replace|Undo|Redo|Statistics|Insert date|Insert time|Select all|Clear selection|Document audit",
 draw: "Shape gallery|Insert picture|Whiteboard|Text box|Callout|Arrow|Line|Rectangle|Circle|Diamond|Star|Flowchart process|Flowchart decision|Shape fill|Shape outline|Shape width|Shape height|Rotate shape|Opacity|Duplicate object|Bring forward|Send backward|Align left|Align centre|Align right|Group|Ungroup|Add caption|Add alt text|Shape inventory",
 ailab: "AI Assistant|Improve clarity|Make formal|Make concise|Expand text|Summarize|Create outline|Convert to bullets|Fix punctuation|Simplify language|Academic tone|Friendly tone|Professional tone|Generate title|Generate intro|Generate conclusion|Create FAQ|Create study notes|Create checklist|Create email|Create report|Meeting minutes|Agenda|Key points|Explain simply|Continue writing|Rewrite|Repeated words|Document audit|Safe cleanup",
 mailings: "Mail merge guide|Recipient list|Insert name field|Insert address field|Greeting line|Date field|Create labels|Create envelope|Preview merge|Next record|Previous record|First record|Last record|Filter recipients|Sort recipients|Add recipient|Edit recipient|Remove recipient|Import CSV|Export recipient list|Merge current|Merge all|Letter template|Invitation template|Certificate template|Address block|Email draft|Mailing statistics|Merge instructions|Print labels",
 forms: "Text field|Checkbox|Radio options|Dropdown|Date field|Signature line|Initials|Section heading|Form title|Instructions|Required marker|Optional marker|Two-column form|Three-column form|Feedback form|Registration form|Survey form|Quiz form|Application form|Contact form|Consent form|Clear fields|Duplicate field|Delete field|Move field up|Move field down|Preview form|Print form|Export form|Form checklist",
 data: "Insert table|Sort A-Z|Sort Z-A|Add row|Delete row|Add column|Delete column|Table header|Blue table|Green table|Dark table|Minimal table|Convert text to table|Convert table to text|Sort paragraphs|Remove blank lines|Trim spaces|Find repeated words|Word count|Character count|Paragraph count|Heading count|Link inventory|Image inventory|Table inventory|Shape inventory|Insert checklist|Insert data grid|Export text|Refresh statistics",
 security: "Privacy overview|Local storage info|Clear recent list|Clear autosave pointer|Export backup|Remove comments|Remove hyperlinks|Remove images|Strip formatting|Plain text export|Show metadata|Read-only copy|Integrity overview|List links|Audit image alt text|Audit external links|Local-only mode|Backup reminder|Security checklist|Clear temporary marks|Review comments|Accept comment|Reject comment|Clean HTML|Safe link guidance|Document copy|Print privacy note|Data handling guide|Security report|Privacy checklist",
 accessibility: "Add image alt text|Heading audit|Heading 1|Heading 2|Heading 3|Set language|Improve link text|Table header|Table audit|Readable font|High contrast|Increase text size|Left align|Colour-only audit|Document title|Document summary|Empty heading audit|Vague link audit|Image captions|List structure audit|Reading order|Accessibility checklist|Plain language|Form instructions|Label form field|Chart description|Add page heading|Accessible print|Check headings|Accessibility report",
 developer: "Show HTML source|Insert HTML|Insert CSS class|Inspect selection|Selection HTML|Selection text|DOM count|List headings|List images|List tables|List links|List shapes|Clean HTML|Format HTML|Escape selected text|Copy selected HTML|Code block|Inline code|Preformatted text|Divider|Semantic section|Article block|Details block|Quote block|Abbreviation|Keyboard key|Variable|Sample output|Developer notes|HTML audit",
 automate: "Auto style headings|Auto create TOC|Auto create outline|Auto bullet paragraphs|Auto number paragraphs|Auto clean spaces|Auto remove blank lines|Title case headings|Sentence case|Auto page breaks|Style tables|Alt text audit|Section dividers|Format quotes|Format code|Create checklist|Insert date|Insert page number|Report skeleton|Meeting notes|Study notes|Summary|FAQ|Email draft|Letter draft|Agenda|Minutes|Document audit|Safe cleanup|Run all cleanup"
};
Object.keys(PRO_FEATURES).forEach(k=>PRO_FEATURES[k]=PRO_FEATURES[k].split('|'));
function proRibbonGroup(tab) {
 const names=PRO_FEATURES[tab]||[];
 return `<div class="rgroup"><div><div class="rgroup-inner" style="max-width:1100px">${names.map((n,i)=>`<div class="rtool"><button title="${escapeHTML(n)}" onclick="runProFeature('${tab}',${i})">${['✦','▣','Aa','◇','✓','↗','☷','✎'][i%8]}</button><small>${escapeHTML(n)}</small></div>`).join('')}</div><div class="rlabel">${tab.toUpperCase()} · ${names.length} tools</div></div></div>`;
}
function runProFeature(tab,i) {
 const n=PRO_FEATURES[tab]?.[i]; if(!n)return;
 const direct={
 'New':()=>createBlank(),'Save':()=>saveDocument(),'Export HTML':()=>exportHTML(),'Export TXT':()=>exportTXT(),'Print PDF':()=>requestPrint(),
 'Open':()=>document.getElementById('openFile').click(),'Backup':()=>exportHTML(),'Autosave':()=>saveDocument(),
 'Cover page':()=>insertCoverPage(),'Page break':()=>insertPageBreak(),'Header':()=>editHeader(),'Footer':()=>editFooter(),
 'Page number':()=>insertPageNumber(),'TOC':()=>insertTOC(),'Bibliography':()=>insertBibliography(),'Footnote':()=>insertFootnote(),
 'Citation':()=>insertCitation(),'Find':()=>openFind(),'Replace':()=>openFind(),'Undo':()=>doFormat('undo'),'Redo':()=>doFormat('redo'),
 'Statistics':()=>showStatistics(),'Insert date':()=>insertDate(),'Insert time':()=>insertHTML(new Date().toLocaleTimeString()),
 'Select all':()=>doFormat('selectAll'),'Clear selection':()=>getSelection().removeAllRanges(),
 'Shape gallery':()=>showShapeGallery(),'Insert picture':()=>chooseImage(),'Whiteboard':()=>openWhiteboard(),
 'Text box':()=>insertTextBox(),'Callout':()=>insertCallout(),'Arrow':()=>insertShapePresetByName('Arrow right'),
 'Line':()=>insertHTML('<hr>'),'Rectangle':()=>insertShapePresetByName('Rectangle'),'Circle':()=>insertShapePresetByName('Circle'),
 'Diamond':()=>insertShapePresetByName('Diamond'),'Star':()=>insertShapePresetByName('Star 5 point'),
 'Flowchart process':()=>insertShapePresetByName('Flowchart process'),'Flowchart decision':()=>insertShapePresetByName('Flowchart decision'),
 'AI Assistant':()=>openWritingAssistant(),'Improve clarity':()=>runAIQuick('clear'),'Make formal':()=>runAIQuick('formal'),
 'Make concise':()=>runAIQuick('concise'),'Expand text':()=>runAIQuick('expand'),'Summarize':()=>runAIQuick('summary'),
 'Create outline':()=>runAIQuick('outline'),'Convert to bullets':()=>runAIQuick('bullets'),'Fix punctuation':()=>runAIQuick('grammar'),
 'Simplify language':()=>runAIQuick('simple'),'Academic tone':()=>runAIQuick('academic'),'Friendly tone':()=>runAIQuick('friendly'),
 'Professional tone':()=>runAIQuick('formal'),'Generate title':()=>runAIQuick('title'),'Generate intro':()=>runAIQuick('introduction'),
 'Generate conclusion':()=>runAIQuick('conclusion'),'Create FAQ':()=>runAIQuick('faq'),'Create study notes':()=>runAIQuick('study'),
 'Create checklist':()=>insertChecklist(),'Create email':()=>runAIQuick('email'),'Create report':()=>runAIQuick('report'),
 'Meeting minutes':()=>runAIQuick('minutes'),'Agenda':()=>runAIQuick('agenda'),'Key points':()=>runAIQuick('keypoints'),
 'Explain simply':()=>runAIQuick('explain'),'Continue writing':()=>runAIQuick('continue'),'Rewrite':()=>runAIQuick('rewrite'),
 'Repeated words':()=>findRepeatedWords(),'Document audit':()=>runAIQuick('audit'),'Safe cleanup':()=>cleanupDocument(),
 'Insert table':()=>insertTable(),'Blue table':()=>styleTable('blue'),'Green table':()=>styleTable('green'),
 'Dark table':()=>styleTable('dark'),'Minimal table':()=>styleTable('minimal'),'Insert checklist':()=>insertChecklist(),
 'Word count':()=>showStatistics(),'Character count':()=>showStatistics(),'Paragraph count':()=>showStatistics(),
 'Heading 1':()=>applyBlock('h1'),'Heading 2':()=>applyBlock('h2'),'Heading 3':()=>applyBlock('h3'),
 'Readable font':()=>setFont('Arial'),'High contrast':()=>document.body.classList.toggle('dark'),
 'Increase text size':()=>setFontSize(18),'Left align':()=>doFormat('justifyLeft'),
 'Auto style headings':()=>autoStyleHeadings(),'Auto create TOC':()=>insertTOC(),'Auto create outline':()=>runAIQuick('outline'),
 'Auto bullet paragraphs':()=>runAIQuick('bullets'),'Auto clean spaces':()=>cleanupDocument(),
 'Auto remove blank lines':()=>removeBlankParagraphs(),'Title case headings':()=>titleCaseHeadings(),
 'Sentence case':()=>runAIQuick('sentenceCase'),'Auto page breaks':()=>insertPageBreak(),
 'Style tables':()=>styleTable('blue'),'Section dividers':()=>insertHTML('<hr>'),'Format quotes':()=>applyBlock('blockquote'),
 'Format code':()=>insertHTML('<pre>Code goes here</pre>'),'Insert page number':()=>insertPageNumber(),
 'Summary':()=>runAIQuick('summary'),'FAQ':()=>runAIQuick('faq'),'Email draft':()=>runAIQuick('email'),
 'Letter draft':()=>insertHTML('<h2>Letter</h2><p>Dear __________,</p><p></p><p>Regards,</p>'),
 'Run all cleanup':()=>cleanupDocument()
 };
 if(direct[n]){direct[n]();return;}
 if(n==='Duplicate'){document.getElementById('docTitle').value+=' copy';saveDocument();return;}
 if(n==='Rename'){const v=prompt('Document title:',document.getElementById('docTitle').value);if(v!==null)document.getElementById('docTitle').value=v;return;}
 if(n==='Properties'){showProInventory('Document properties',[`Title: ${document.getElementById('docTitle').value}`,`Words: ${page.innerText.trim().split(/\\s+/).filter(Boolean).length}`,`Characters: ${page.innerText.length}`]);return;}
 if(n==='Remove hyperlinks'){page.querySelectorAll('a').forEach(x=>x.replaceWith(...x.childNodes));toast('Hyperlinks removed');return;}
 if(n==='Remove images'){page.querySelectorAll('img').forEach(x=>x.remove());toast('Images removed');return;}
 if(n==='Strip formatting'){page.querySelectorAll('*').forEach(x=>x.removeAttribute('style'));toast('Inline formatting removed');return;}
 if(n==='Remove comments'){page.querySelectorAll('.kms-comment').forEach(x=>x.remove());toast('Comments removed');return;}
 if(n==='Clear recent list'){try{localStorage.removeItem(STORAGE_KEY);renderRecent();toast('Recent list cleared')}catch(e){toast('Storage unavailable')}return;}
 if(n==='Clear autosave pointer'){try{localStorage.removeItem(ACTIVE_KEY);toast('Autosave pointer cleared')}catch(e){toast('Storage unavailable')}return;}
 if(n==='Sort paragraphs'){sortProParagraphs();return;}
 if(n==='Remove blank lines'){removeBlankParagraphs();return;}
 if(n==='Trim spaces'){cleanupDocument();return;}
 if(n==='Repeated words'){findRepeatedWords();return;}
 if(n==='Show HTML source'){showProInventory('HTML source',[page.innerHTML]);return;}
 if(n==='DOM count'){toast('Elements: '+page.querySelectorAll('*').length);return;}
 if(n==='List headings'||n==='Heading audit'||n==='Check headings'){showProInventory(n,[...page.querySelectorAll('h1,h2,h3')].map(x=>x.tagName+' — '+x.innerText));return;}
 if(n==='List images'||n==='Audit image alt text'||n==='Alt text audit'){showProInventory(n,[...page.querySelectorAll('img')].map(x=>x.alt||'MISSING ALT TEXT'));return;}
 if(n==='List tables'){showProInventory(n,[...page.querySelectorAll('table')].map(x=>x.rows.length+' rows'));return;}
 if(n==='List links'){showProInventory(n,[...page.querySelectorAll('a')].map(x=>x.href));return;}
 if(n==='List shapes'||n==='Shape inventory'){showProInventory(n,[...page.querySelectorAll('.kms-shape')].map(x=>x.dataset.shape||'Shape'));return;}
 if(n==='Add image alt text'){const im=page.querySelector('img.object-selected')||page.querySelector('img');if(!im){toast('Insert an image first');return;}const v=prompt('Alternative text:',im.alt||'');if(v!==null)im.alt=v;return;}
 if(n==='Insert name field'||n==='Insert address field'||n==='Greeting line'||n==='Date field'||n==='Address block'){insertHTML(`<p>${escapeHTML(n)}: __________________________</p>`);return;}
 if(n==='Checkbox'){insertHTML('<p>☐ Checkbox option</p>');return;}
 if(n==='Text field'||n==='Dropdown'||n==='Radio options'||n==='Signature line'||n==='Initials'){insertHTML(`<p>${escapeHTML(n)}: __________________________</p>`);return;}
 if(n==='Create labels'||n==='Create envelope'||n==='Letter template'||n==='Invitation template'||n==='Certificate template'||n==='Email draft'){runAIQuick('email');return;}
 if(n==='Insert HTML'){const v=prompt('HTML snippet to insert:','<p>New section</p>');if(v!==null)insertHTML(v);return;}
 if(n==='Selection text'||n==='Selection HTML'||n==='Copy selected HTML'){const s=getSelection().toString();if(s)navigator.clipboard?.writeText(s);else toast('Select text first');return;}
 if(n==='Clean HTML'||n==='Safe cleanup'||n==='Auto clean spaces'){cleanupDocument();return;}
 if(n==='Insert CSS class'){const v=prompt('CSS class name:');if(v)insertHTML(`<div class="${escapeHTML(v)}">Click to edit this section</div>`);return;}
 if(n==='Export form'||n==='Export form HTML'){exportHTML();return;}
 if(n==='Print form'){requestPrint();return;}
 if(n==='Preview form'){showDialog('<h2>Form preview</h2><p>Use Print or Export HTML to preview this form.</p><div class="dialog-actions"><button onclick="closeDialog()">Close</button></div>');return;}
 if(n==='Remove blank lines'){removeBlankParagraphs();return;}
 if(n==='Sort A-Z'||n==='Sort Z-A'){sortProTable(n==='Sort A-Z');return;}
 if(n==='Add row'||n==='Delete row'||n==='Add column'||n==='Delete column'){editProTable(n);return;}
 if(n==='Table header'){const t=page.querySelector('table');if(!t){toast('Insert a table first');return;}[...t.rows[0].cells].forEach(c=>{const x=document.createElement('th');x.innerHTML=c.innerHTML;c.replaceWith(x)});return;}
 if(n==='Convert text to table'){convertProTextToTable();return;}
 if(n==='Convert table to text'){const t=page.querySelector('table');if(!t){toast('Insert a table first');return;}const x=document.createElement('pre');x.textContent=[...t.rows].map(r=>[...r.cells].map(c=>c.innerText).join('\\t')).join('\\n');t.replaceWith(x);return;}
 if(n==='Set language'||n==='Document title'||n==='Document summary'||n==='Form title'||n==='Section heading'||n==='Instructions'||n==='Required marker'||n==='Optional marker'||n==='Add caption'||n==='Chart description'||n==='Improve link text'||n==='Semantic section'||n==='Article block'||n==='Details block'||n==='Quote block'||n==='Abbreviation'||n==='Keyboard key'||n==='Variable'||n==='Sample output'||n==='Developer notes'||n==='Recipient list'||n==='Mail merge guide'||n==='Privacy overview'||n==='Local storage info'||n==='Security checklist'||n==='Accessibility checklist'||n==='Accessibility report'||n==='Data handling guide'||n==='Security report'||n==='Backup reminder'||n==='Local-only mode'||n==='Integrity overview'||n==='Safe link guidance'||n==='Print privacy note'||n==='Security report'||n==='Mailing statistics'||n==='Merge instructions'||n==='Filter recipients'||n==='Sort recipients'||n==='Import CSV'||n==='Export recipient list'||n==='Merge current'||n==='Merge all'||n==='Next record'||n==='Previous record'||n==='First record'||n==='Last record'||n==='Add recipient'||n==='Edit recipient'||n==='Remove recipient'||n==='Preview merge'||n==='Feedback form'||n==='Registration form'||n==='Survey form'||n==='Quiz form'||n==='Application form'||n==='Contact form'||n==='Consent form'||n==='Two-column form'||n==='Three-column form'||n==='Form checklist'||n==='Clear fields'||n==='Duplicate field'||n==='Delete field'||n==='Move field up'||n==='Move field down'||n==='Heading audit'||n==='Table audit'||n==='Colour-only audit'||n==='Empty heading audit'||n==='Vague link audit'||n==='Image captions'||n==='List structure audit'||n==='Reading order'||n==='Plain language'||n==='Form instructions'||n==='Label form field'||n==='Accessible print'||n==='Accessibility report'||n==='HTML audit'||n==='Format HTML'||n==='Escape selected text'||n==='Code block'||n==='Inline code'||n==='Preformatted text'||n==='Divider'||n==='Print labels'||n==='Create data grid'){
  const v=prompt(n+': enter content/details (Cancel to skip).',n);if(v!==null)insertHTML(`<p>${escapeHTML(v)}</p>`);return;
 }
 if(n==='Privacy checklist'||n==='Security report'||n==='Security checklist'||n==='Accessibility checklist'||n==='Accessibility report'){showDialog(`<h2>${escapeHTML(n)}</h2><p>This edition stores documents locally in the Qt WebEngine browser profile. It does not include cloud collaboration, password encryption or a connected online AI provider. Keep backups and add descriptive image alt text.</p><div class="dialog-actions"><button onclick="closeDialog()">Close</button></div>`);return;}
 const v=prompt(n+': enter details to insert (Cancel to skip).',n);if(v!==null)insertHTML(`<p>${escapeHTML(v)}</p>`);
}
function showProInventory(title,items){showDialog(`<h2>${escapeHTML(title)}</h2>${items.length?'<ul>'+items.map(x=>'<li>'+escapeHTML(x)+'</li>').join('')+'</ul>':'<p>No items found.</p>'}<div class="dialog-actions"><button onclick="closeDialog()">Close</button></div>`);}
function insertPageNumber(){insertHTML('<span class="page-number">Page 1</span>');}
function insertShapePresetByName(n){const i=SHAPE_PRESETS.findIndex(x=>x[0].toLowerCase()===n.toLowerCase());if(i>=0)insertShapePreset(i);else showShapeGallery();}
function duplicateSelectedObject(){if(!selectedObject){toast('Select an object first');return;}const x=selectedObject.cloneNode(true);selectedObject.after(x);selectDocumentObject(x);}
function moveSelectedObject(d){if(!selectedObject){toast('Select an object first');return;}if(d>0)selectedObject.parentNode.appendChild(selectedObject);else selectedObject.parentNode.insertBefore(selectedObject,selectedObject.parentNode.firstChild);}
function alignSelected(a){if(!selectedObject){toast('Select an object first');return;}selectedObject.style.display='block';selectedObject.style.margin=a==='center'?'8px auto':'8px 0';if(a==='right')selectedObject.style.marginLeft='auto';}
function editSelectedShape(p){if(!selectedObject){toast('Select a shape first');return;}const prop={'Shape fill':'backgroundColor','Shape outline':'borderColor','Shape width':'width','Shape height':'height','Rotate shape':'transform','Opacity':'opacity'}[p]||p;const v=prompt('New '+p+':',selectedObject.style[prop]||'');if(v===null)return;if(prop==='transform')selectedObject.style.transform=`rotate(${parseFloat(v)||0}deg)`;else if(['width','height'].includes(prop))selectedObject.style[prop]=/px$/.test(v)?v:v+'px';else selectedObject.style[prop]=v;}
function sortProParagraphs(){const a=[...page.querySelectorAll(':scope > p')];a.sort((x,y)=>x.innerText.localeCompare(y.innerText));a.forEach(x=>page.appendChild(x));toast('Paragraphs sorted A–Z');}
function removeBlankParagraphs(){page.querySelectorAll('p').forEach(x=>{if(!x.textContent.trim()&&!x.querySelector('*'))x.remove();});updateStats();toast('Blank paragraphs removed');}
function cleanupDocument(){page.innerHTML=page.innerHTML.replace(/&nbsp;/g,' ').replace(/[ \\t]{2,}/g,' ');removeBlankParagraphs();}
function titleCaseHeadings(){page.querySelectorAll('h1,h2,h3').forEach(x=>x.textContent=x.textContent.toLowerCase().replace(/\\b\\w/g,c=>c.toUpperCase()));toast('Headings converted to title case');}
function runAIQuick(mode){openWritingAssistant();const el=document.getElementById('aiInput');if(el)el.value=page.innerText.trim();runLocalAI(mode);}
function findRepeatedWords(){const w=page.innerText.toLowerCase().match(/\\b[\\w']+\\b/g)||[],c={};w.forEach(x=>c[x]=(c[x]||0)+1);showProInventory('Repeated words',Object.entries(c).filter(x=>x[1]>2).map(x=>x[0]+': '+x[1]));}
function autoStyleHeadings(){page.querySelectorAll('h1,h2,h3').forEach((x,i)=>{x.style.fontFamily='Arial';x.style.color=i?'#245bb5':'#173f70';});toast('Heading styles applied');}
function insertCallout(){insertHTML('<div class="callout">Important note — click to edit</div>');}
function insertChecklist(){insertHTML('<ul style="list-style:none"><li>☐ Task one</li><li>☐ Task two</li><li>☐ Task three</li></ul>');}
function sortProTable(asc){const t=page.querySelector('table');if(!t||t.rows.length<3){toast('Insert a table with at least two data rows');return;}const b=t.tBodies[0]||t,rows=[...b.rows].slice(1).sort((x,y)=>x.innerText.localeCompare(y.innerText)*(asc?1:-1));rows.forEach(x=>b.appendChild(x));}
function editProTable(a){const t=page.querySelector('table');if(!t){toast('Insert a table first');return;}const r=t.rows[t.rows.length-1];if(a==='Add row'){const x=t.insertRow();for(let i=0;i<r.cells.length;i++)x.insertCell().textContent='New cell';}else if(a==='Delete row'){if(t.rows.length>1)t.deleteRow(-1);}else if(a==='Add column'){[...t.rows].forEach(x=>x.insertCell().textContent='New cell');}else{[...t.rows].forEach(x=>x.deleteCell(-1));}}
function convertProTextToTable(){const v=prompt('Paste rows (new line) and columns (comma):');if(v===null)return;insertHTML('<table><tbody>'+v.split(/\\r?\\n/).filter(Boolean).map((r,i)=>'<tr>'+r.split(',').map(c=>i===0?'<th>'+escapeHTML(c)+'</th>':'<td>'+escapeHTML(c)+'</td>').join('')+'</tr>').join('')+'</tbody></table>');}

// 1,000 font style combinations, 400 text styles, 100 bullet markers, 400 shape variants.
const FONT_STYLE_PRESETS=Array.from({length:1000},(_,i)=>{const f=[...new Set(FONT_FAMILIES)],c=['#202938','#173f70','#245bb5','#16834c','#7652b8','#9d174d','#0f766e','#9a5b13'];return {family:f[i%f.length],size:[12,14,16,18,20,24,28,32][Math.floor(i/7)%8],weight:[400,500,600,700,800,900][Math.floor(i/11)%6],italic:i%3===0,color:c[Math.floor(i/13)%8],spacing:(i%5)*.2};});
const PRO_TEXT_STYLES=Array.from({length:400},(_,i)=>{const c=['#173f70','#245bb5','#16834c','#7652b8','#b45309','#9d174d','#0f766e','#334155'];return {name:'Style '+(i+1),css:`font-size:${[16,18,20,22,24,26,28,30,32][i%9]}px;font-weight:${[500,600,700,800,900][Math.floor(i/3)%5]};color:${c[Math.floor(i/5)%8]};letter-spacing:${(i%5)*.2}px`};});
const PRO_BULLETS=Array.from({length:100},(_,i)=>['•','○','▪','◆','➤','➜','✓','✔','☑','★','☆','✦','❖','➢','»','→','♥','●','–','☀','✧','◉','▸','◦','✪','✳','✴','✷','❯','☞'][i%30]);
const PRO_SHAPES=Array.from({length:400},(_,i)=>{const s=SHAPE_PRESETS[i%SHAPE_PRESETS.length];return {name:s[0]+' Style '+(Math.floor(i/SHAPE_PRESETS.length)+1),glyph:s[1],base:s[0]};});
function showFontStyleGallery(){showDialog(`<h2>1,000 font style samples</h2><p>1,000 combinations of font family, size, weight, colour and spacing. Actual typefaces depend on fonts installed on this PC.</p><div class="gallery-grid">${FONT_STYLE_PRESETS.map((x,i)=>`<button class="gallery-item" onclick="applyFontStyle(${i})"><span class="gallery-preview" style="font-family:'${escapeHTML(x.family)}';font-size:${Math.min(x.size,28)}px;font-weight:${x.weight};font-style:${x.italic?'italic':'normal'};color:${x.color};letter-spacing:${x.spacing}px">Aa Sample</span><small>${i+1}. ${escapeHTML(x.family)}<br>${x.weight} weight</small></button>`).join('')}</div><div class="dialog-actions"><button onclick="closeDialog()">Close</button></div>`);}
function applyFontStyle(i){const x=FONT_STYLE_PRESETS[i];if(!x)return;const s=getSelection(),v=s&&s.rangeCount&&!s.isCollapsed?escapeHTML(s.toString()):'Sample text';insertHTML(`<span style="font-family:'${escapeHTML(x.family)}';font-size:${x.size}px;font-weight:${x.weight};font-style:${x.italic?'italic':'normal'};color:${x.color};letter-spacing:${x.spacing}px">${v}</span>`);}
function showTextStyles(){showDialog(`<h2>400 text styles</h2><p>Each tile shows a live sample. Select text before applying.</p><div class="gallery-grid">${PRO_TEXT_STYLES.map((x,i)=>`<button class="gallery-item" onclick="applyProTextStyle(${i})"><span class="gallery-preview" style="${x.css}">Aa Sample</span><small>${i+1}. ${x.name}</small></button>`).join('')}</div><div class="dialog-actions"><button onclick="closeDialog()">Close</button></div>`);}
function applyProTextStyle(i){const x=PRO_TEXT_STYLES[i];if(!x)return;const s=getSelection(),v=s&&s.rangeCount&&!s.isCollapsed?escapeHTML(s.toString()):'Sample heading';insertHTML(`<span style="${x.css}">${v}</span>`);}
function showBulletGallery(){showDialog(`<h2>100 bullet styles</h2><div class="gallery-grid">${PRO_BULLETS.map((b,i)=>`<button class="gallery-item" onclick="applyProBullet(${i})"><span class="gallery-preview">${b} Sample item</span><small>${i+1}. Bullet style</small></button>`).join('')}</div><div class="dialog-actions"><button onclick="closeDialog()">Close</button></div>`);}
function applyProBullet(i){const b=PRO_BULLETS[i],v=prompt('List items, one per line:','First item\\nSecond item\\nThird item');if(v===null)return;insertHTML('<ul style="list-style:none">'+v.split(/\\r?\\n/).filter(x=>x.trim()).map(x=>`<li>${b} ${escapeHTML(x)}</li>`).join('')+'</ul>');}
function showShapeGallery(){showDialog(`<h2>400 shape variations</h2><p>Editable shape labels; select an inserted shape to adjust its object properties.</p><div class="gallery-grid">${PRO_SHAPES.map((s,i)=>`<button class="gallery-item" onclick="insertProShape(${i})"><span class="gallery-preview">${s.glyph}</span><small>${i+1}. ${escapeHTML(s.name)}</small></button>`).join('')}</div><div class="dialog-actions"><button onclick="closeDialog()">Close</button></div>`);}
function insertProShape(i){const s=PRO_SHAPES[i],v=prompt('Shape text:',s.name);if(v===null)return;insertHTML(`<span class="kms-shape" data-shape="${escapeHTML(s.base)}" contenteditable="true" style="background:#dbeafe;color:#174b91;border:2px solid #245bb5;width:150px;height:88px">${escapeHTML(v)}</span>`);}
const _setTabV20=setTab;
setTab=function(t){_setTabV20(t);if(t==='home'){const g=document.querySelector('#ribbon .rgroup:nth-child(2) .rgroup-inner');if(g&&!g.querySelector('[data-fontgallery]')){const b=document.createElement('button');b.dataset.fontgallery='1';b.textContent='1,000 Font Samples';b.onclick=showFontStyleGallery;g.appendChild(b);}}};

// Local text-processing assistant. It does not pretend to call a remote generative AI.
function runLocalAI(mode){
 const el=document.getElementById('aiInput');if(!el)return;const v=el.value.trim();if(!v){toast('Enter text first');return;}
 const ss=v.match(/[^.!?]+[.!?]*/g)||[v];let o=v;
 if(mode==='formal')o=v.replace(/\\bcan't\\b/gi,'cannot').replace(/\\bdon't\\b/gi,'do not').replace(/\\bwon't\\b/gi,'will not').replace(/\\bkids\\b/gi,'children');
 else if(mode==='clear'||mode==='simple')o=v.replace(/\\butilize\\b/gi,'use').replace(/\\bin order to\\b/gi,'to').replace(/\\bdue to the fact that\\b/gi,'because');
 else if(mode==='concise')o=ss.slice(0,Math.max(1,Math.ceil(ss.length*.65))).join(' ').trim();
 else if(mode==='expand')o=v+'\\n\\nFurther explanation: '+ss.map(x=>x.trim()).join(' ');
 else if(mode==='summary')o=ss.slice(0,Math.max(1,Math.ceil(ss.length/3))).join(' ').trim();
 else if(mode==='grammar')o=v.replace(/\\s+([,.!?;:])/g,'$1').replace(/([,.!?])(?=[A-Za-z])/g,'$1 ').replace(/\\bi\\b/g,'I').replace(/\\s{2,}/g,' ').trim();
 else if(['bullets','keypoints','checklist'].includes(mode))o=ss.map(x=>x.trim()).filter(Boolean).map(x=>(mode==='checklist'?'☐ ':'• ')+x).join('\\n');
 else if(['outline','study'].includes(mode))o='Overview\\n'+ss.map((x,i)=>(i+1)+'. '+x.trim()).join('\\n');
 else if(mode==='title')o=ss[0].trim().split(/\\s+/).slice(0,9).join(' ').replace(/^\\w/,c=>c.toUpperCase());
 else if(mode==='introduction')o='Introduction\\n'+ss.slice(0,2).join(' ').trim();
 else if(mode==='conclusion')o='In conclusion, '+ss.slice(-2).join(' ').trim();
 else if(mode==='friendly')o=v.replace(/\\btherefore\\b/gi,'so').replace(/\\bfurthermore\\b/gi,'also');
 else if(mode==='academic')o=v.replace(/\\ba lot of\\b/gi,'a substantial amount of').replace(/\\bshows\\b/gi,'demonstrates');
 else if(mode==='explain')o='Key idea:\\n'+v+'\\n\\nIn simple terms: '+ss[0].trim();
 else if(mode==='faq')o='Frequently asked questions\\n'+ss.map((x,i)=>`Q${i+1}. What does this mean?\\nA. ${x.trim()}`).join('\\n\\n');
 else if(mode==='email')o='Subject: '+ss[0].trim().slice(0,60)+'\\n\\nHello,\\n\\n'+v+'\\n\\nRegards,';
 else if(mode==='report')o='Report\\n\\nOverview\\n'+v+'\\n\\nKey points\\n'+ss.map(x=>'• '+x.trim()).join('\\n')+'\\n\\nConclusion\\n'+ss.slice(-1).join(' ');
 else if(['minutes','meeting','agenda'].includes(mode))o=mode.toUpperCase()+'\\n\\n'+ss.map((x,i)=>(i+1)+'. '+x.trim()).join('\\n');
 else if(mode==='repetition'){const w=v.toLowerCase().match(/\\b[\\w']+\\b/g)||[],c={};w.forEach(x=>c[x]=(c[x]||0)+1);o=Object.entries(c).filter(x=>x[1]>1).map(x=>x[0]+': '+x[1]).join('\\n')||'No repeated words found.';}
 else if(mode==='cleanup')o=v.replace(/[ \\t]{2,}/g,' ').replace(/\\n{3,}/g,'\\n\\n').trim();
 else if(mode==='titleCase')o=v.replace(/\\b\\w/g,c=>c.toUpperCase());
 else if(mode==='sentenceCase')o=v.toLowerCase().replace(/(^\\s*\\w|[.!?]\\s+\\w)/g,c=>c.toUpperCase());
 else if(mode==='audit')o=`Words: ${(v.match(/\\b[\\w']+\\b/g)||[]).length}\\nCharacters: ${v.length}\\nSentences: ${ss.length}`;
 else if(mode==='translate')o='Translation requires an online language service, which is not connected in this offline edition.\\n\\n'+v;
 document.getElementById('aiOutput').value=o;
}

</script>
</body>
</html>
"""



# ---------------- ONLINE LICENSING ----------------
APP_NAME = "KMS WORD AI V20 PRO"
APP_DIR = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "KMSWordAI"
APP_DIR.mkdir(parents=True, exist_ok=True)
LICENSE_CACHE = APP_DIR / "activation.json"

def app_resource_path(filename):
    """Resolve bundled files in both source and PyInstaller builds."""
    root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return root / filename

def load_activation_config():
    config_path = app_resource_path("kms_config.json")
    try:
        data = json.loads(config_path.read_text(encoding="utf-8"))
        return data
    except Exception:
        return {"api_base_url": "https://YOUR-ACTIVATION-SERVER.example.com"}

def get_device_id():
    # Windows MachineGuid is stable across ordinary app reinstalls.
    if sys.platform == "win32":
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                 r"SOFTWARE\Microsoft\Cryptography")
            value, _ = winreg.QueryValueEx(key, "MachineGuid")
            raw = "kms-word-ai:" + str(value)
        except Exception:
            raw = "kms-word-ai:" + str(uuid.getnode())
    else:
        raw = "kms-word-ai:" + str(uuid.getnode())
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()

def licensing_request(endpoint, payload):
    config = load_activation_config()
    base = str(config.get("api_base_url", "")).rstrip("/")
    if not base or "YOUR-ACTIVATION-SERVER" in base:
        raise RuntimeError("The activation server URL has not been configured.")
    request = urllib.request.Request(
        base + endpoint,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "User-Agent": "KMSWordAI/20"},
        method="POST"
    )
    try:
        with urllib.request.urlopen(request, timeout=12) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        try:
            detail = json.loads(exc.read().decode("utf-8")).get("detail", "Activation request rejected.")
        except Exception:
            detail = "Activation request rejected."
        raise RuntimeError(str(detail))
    except Exception as exc:
        raise RuntimeError("Could not contact the activation server. Check your internet connection. " + str(exc))

def validate_saved_activation():
    try:
        saved = json.loads(LICENSE_CACHE.read_text(encoding="utf-8"))
        token = saved.get("token", "")
        if not token:
            return False
        result = licensing_request("/validate", {
            "token": token,
            "device_id": get_device_id()
        })
        return bool(result.get("valid"))
    except Exception:
        return False

class ActivationDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Activate KMS WORD AI V20 PRO")
        self.setFixedWidth(470)
        self.setWindowFlag(self.windowFlags() & ~__import__("PyQt6").QtCore.Qt.WindowType.WindowCloseButtonHint)
        self.setStyleSheet("""
            QDialog { background:#f3f6fb; }
            QLabel#heading { color:#14243d; font-size:22px; font-weight:800; }
            QLabel { color:#35445b; font-size:13px; }
            QLineEdit { padding:12px; border:1px solid #c8d4e4; border-radius:7px; background:white; font-size:15px; }
            QPushButton { padding:11px; border-radius:7px; font-weight:700; background:#245bb5; color:white; }
            QPushButton:disabled { background:#9aa9bd; }
            QLabel#status { color:#a33; }
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 26, 28, 24)
        layout.setSpacing(13)
        heading = QLabel("KMS WORD AI")
        heading.setObjectName("heading")
        layout.addWidget(heading)
        sub = QLabel("Version 20.0 PRO · Product activation")
        layout.addWidget(sub)
        layout.addSpacing(8)
        layout.addWidget(QLabel("Enter the product key supplied with your purchase."))
        self.key_input = QLineEdit()
        self.key_input.setPlaceholderText("XXXX-XXXX-XXXX-XXXX")
        self.key_input.setMaxLength(40)
        self.key_input.returnPressed.connect(self.activate)
        layout.addWidget(self.key_input)
        self.status = QLabel("Internet connection is required for activation.")
        self.status.setObjectName("status")
        self.status.setWordWrap(True)
        layout.addWidget(self.status)
        self.activate_btn = QPushButton("Activate product")
        self.activate_btn.clicked.connect(self.activate)
        layout.addWidget(self.activate_btn)
        note = QLabel("Your licence is associated with this computer. Contact the software provider if you need to move it.")
        note.setWordWrap(True)
        layout.addWidget(note)
        self.setModal(True)

    def activate(self):
        key = self.key_input.text().strip().upper()
        if not key:
            self.status.setText("Enter your product key.")
            return
        self.activate_btn.setEnabled(False)
        self.status.setText("Checking product key with the activation server…")
        QApplication.processEvents()
        try:
            result = licensing_request("/activate", {
                "license_key": key,
                "device_id": get_device_id(),
                "app_version": "20.0"
            })
            token = result.get("token")
            if not token:
                raise RuntimeError("The server did not return an activation token.")
            LICENSE_CACHE.write_text(json.dumps({
                "token": token,
                "activated_at": result.get("activated_at", "")
            }), encoding="utf-8")
            self.status.setText("Activation successful.")
            self.accept()
        except Exception as exc:
            self.status.setText(str(exc))
            self.activate_btn.setEnabled(True)


class NativeBridge(QObject):
    def __init__(self, window):
        super().__init__(window)
        self.window = window

    @pyqtSlot()
    def printDocument(self):
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        dialog = QPrintDialog(printer, self.window)
        dialog.setWindowTitle("Print KMS WORD AI document")
        if dialog.exec() != QPrintDialog.DialogCode.Accepted:
            return
        try:
            self.window.browser.page().print(printer, lambda ok: self.window.statusBar().showMessage(
                "Document sent to printer" if ok else "Printing failed — check the selected printer", 7000
            ))
        except Exception as exc:
            QMessageBox.critical(self.window, "Print error", f"Could not print the document.\n\n{exc}")


class KMSWordAI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("KMS WORD AI — Version 20.0 PRO")
        self.resize(1500, 950)
        self.browser = QWebEngineView()
        # Ensure local document storage is enabled for autosave and tab startup.
        self.browser.settings().setAttribute(QWebEngineSettings.WebAttribute.LocalStorageEnabled, True)
        self.browser.settings().setAttribute(QWebEngineSettings.WebAttribute.JavascriptEnabled, True)
        self.channel = QWebChannel(self.browser.page())
        self.bridge = NativeBridge(self)
        self.channel.registerObject("kmsBridge", self.bridge)
        self.browser.page().setWebChannel(self.channel)
        self.browser.setHtml(HTML, QUrl("qrc:/"))
        self.setCentralWidget(self.browser)
        self._build_menus()
        self.statusBar().showMessage("KMS WORD AI V20.0 PRO ready — if tabs do not respond, check the app error message")

    def run_js(self, script):
        self.browser.page().runJavaScript(script)

    def _build_menus(self):
        file_menu = self.menuBar().addMenu("&File")
        actions = [
            ("New document", "Ctrl+N", "createBlank()"),
            ("Open document", "Ctrl+O", "document.getElementById('openFile').click()"),
            ("Save", "Ctrl+S", "saveDocument()"),
            ("Export HTML", "", "exportHTML()"),
            ("Export TXT", "", "exportTXT()"),
            ("Print / Save PDF", "Ctrl+P", "requestPrint()"),
        ]
        for label, shortcut, js in actions:
            action = QAction(label, self)
            if shortcut:
                action.setShortcut(QKeySequence(shortcut))
            action.triggered.connect(lambda checked=False, code=js: self.run_js(code))
            file_menu.addAction(action)

        view_menu = self.menuBar().addMenu("&View")
        for label, js in [
            ("Dashboard", "showDashboard(true)"),
            ("Document editor", "showDashboard(false)"),
            ("Focus mode", "toggleFocus()"),
            ("Whiteboard", "openWhiteboard()"),
            ("Dark workspace", "toggleDark()"),
        ]:
            action = QAction(label, self)
            action.triggered.connect(lambda checked=False, code=js: self.run_js(code))
            view_menu.addAction(action)

        tools_menu = self.menuBar().addMenu("&Tools")
        for label, js in [
            ("Find and replace", "openFind()"),
            ("Document statistics", "showStatistics()"),
            ("Templates", "showTemplates()"),
            ("Word count", "showStatistics()"),
        ]:
            action = QAction(label, self)
            action.triggered.connect(lambda checked=False, code=js: self.run_js(code))
            tools_menu.addAction(action)

        help_menu = self.menuBar().addMenu("&Help")
        for label, js in [
            ("Quick guide", "showGuide()"),
            ("Keyboard shortcuts", "showShortcuts()"),
            ("About KMS WORD AI", "aboutApp()"),
        ]:
            action = QAction(label, self)
            action.triggered.connect(lambda checked=False, code=js: self.run_js(code))
            help_menu.addAction(action)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName("KMS WORD AI V20.0 PRO")
    app.setOrganizationName("KMS A VISION")
    if not validate_saved_activation():
        activation = ActivationDialog()
        if activation.exec() != QDialog.DialogCode.Accepted:
            sys.exit(0)
    window = KMSWordAI()
    window.showMaximized()
    sys.exit(app.exec())
