# ═══════════════════════════════════════════════════════════════
#  Prime Bazar Bot — Dual Engine
#  Telebot (Main Bot) + Pyrogram (Userbot Bridge)
#  All settings via JSON files — zero code changes needed.
# ═══════════════════════════════════════════════════════════════

import asyncio
import io
import json
import os
import re
import threading
import time
import unicodedata
from datetime import datetime, timedelta

import pandas as pd
import requests
import telebot
from flask import Flask, render_template_string, request, jsonify
from PIL import Image, ImageDraw, ImageFont
from telebot import types
from mongo_db import (
    db_load_settings, db_save_settings,
    db_load_texts, db_save_texts,
    db_load_market, db_save_market,
    db_load_coupons, db_save_coupons,
    db_load_trxids, db_save_trxids,
    db_load_pending_deps, db_save_pending_deps,
    db_load_rejected_deps, db_save_rejected_deps,
    db_load_manual_orders, db_save_manual_orders,
    db_load_all_users, db_save_one_user, db_save_all_users,
    db_load_stock, db_save_stock, db_append_stock, db_stock_count, db_delete_stock,
    db_append_sold, db_load_sold,
    db_load_reviews, db_append_review,
)
import uuid

try:
    from pyrogram import Client, filters as pf
    from pyrogram.raw import functions as raw_fn, types as raw_types
    PYRO_OK = True
except ImportError:
    PYRO_OK = False


# ─────────────────────────────────────────────
#  Flask Keep-Alive (port 3000)
# ─────────────────────────────────────────────
flask_app = Flask("")

@flask_app.route("/")
def home():
    return "🤖 Prime Bazar Bot is Online!"

# ─────────────────────────────────────────────────────────────────────────────
#  OTP Extractor — Microsoft Graph API
# ─────────────────────────────────────────────────────────────────────────────
def get_access_token(client_id: str, refresh_token: str) -> str:
    resp = requests.post(
        "https://login.microsoftonline.com/common/oauth2/v2.0/token",
        data={
            "client_id":     client_id,
            "refresh_token": refresh_token,
            "grant_type":    "refresh_token",
            "scope":         "https://graph.microsoft.com/Mail.Read offline_access",
        },
        timeout=15,
    )
    resp.raise_for_status()
    data = resp.json()
    if "access_token" not in data:
        raise ValueError(data.get("error_description", "Token refresh failed."))
    return data["access_token"]


def fetch_email_code(access_token: str) -> dict:
    import re as _re2
    headers = {"Authorization": f"Bearer {access_token}"}
    params  = {
        "$top":     "10",
        "$select":  "subject,bodyPreview,body,receivedDateTime,from",
        "$orderby": "receivedDateTime desc",
    }
    resp = requests.get(
        "https://graph.microsoft.com/v1.0/me/messages",
        headers=headers, params=params, timeout=15,
    )
    resp.raise_for_status()
    messages = resp.json().get("value", [])
    if not messages:
        return {"error": "No emails found in inbox."}
    for msg in messages:
        subject  = msg.get("subject", "") or ""
        received = msg.get("receivedDateTime", "")
        body_txt = msg.get("bodyPreview", "") or ""
        if not body_txt:
            body_txt = msg.get("body", {}).get("content", "") or ""
        sender = (msg.get("from", {}) or {}).get("emailAddress", {}).get("address", "")
        search_text = subject + " " + body_txt
        codes = _re2.findall(r'\b(\d{4,8})\b', search_text)
        if codes:
            return {
                "subject":  subject,
                "received": received,
                "sender":   sender,
                "code":     codes[0],
            }
    return {"error": "No OTP code found in recent emails."}


_OTP_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no"/>
<title>Prime Bazar — 2FA &amp; OTP</title>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap" rel="stylesheet"/>
<script src="https://telegram.org/js/telegram-web-app.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box;-webkit-tap-highlight-color:transparent}
:root{
  --bg:#080b12;
  --surface:#0f1520;
  --surface2:#141b28;
  --border:rgba(255,255,255,0.07);
  --border-active:rgba(99,157,255,0.45);
  --accent:#639dff;
  --accent-dark:#3b6fd4;
  --accent-glow:rgba(99,157,255,0.15);
  --text:#e8edf8;
  --text-secondary:#8a96b0;
  --text-dim:#4a5568;
  --success:#3dd68c;
  --success-bg:rgba(61,214,140,0.1);
  --error:#f06080;
  --error-bg:rgba(240,96,128,0.08);
  --warn:#f5a623;
  --radius:14px;
  --radius-sm:10px;
  --radius-xs:7px;
}
html,body{height:100%;background:var(--bg)}
body{
  font-family:'Inter',sans-serif;color:var(--text);
  min-height:100vh;padding:0 0 60px;
  background:
    radial-gradient(ellipse 70% 40% at 15% 0%,rgba(63,117,255,0.07) 0%,transparent 70%),
    radial-gradient(ellipse 50% 30% at 85% 100%,rgba(99,157,255,0.05) 0%,transparent 60%),
    var(--bg);
}

/* ── HEADER ── */
.header{
  text-align:center;padding:22px 20px 16px;
  border-bottom:1px solid var(--border);
  background:linear-gradient(180deg,rgba(99,157,255,0.04) 0%,transparent 100%);
}
.header-logo{
  width:48px;height:48px;border-radius:14px;margin:0 auto 10px;
  background:linear-gradient(135deg,var(--accent-dark),var(--accent));
  display:flex;align-items:center;justify-content:center;
  font-size:22px;box-shadow:0 4px 20px rgba(99,157,255,0.25);
}
.header h1{font-size:17px;font-weight:800;letter-spacing:-.2px;color:var(--text)}
.header p{font-size:11.5px;color:var(--text-secondary);margin-top:3px}

/* ── TAB BAR ── */
.tabs{
  display:flex;background:var(--surface);
  border-bottom:1px solid var(--border);padding:0 16px;gap:4px;
}
.tab{
  flex:1;padding:12px 8px 10px;font-size:12px;font-weight:600;
  color:var(--text-secondary);border:none;background:none;cursor:pointer;
  border-bottom:2px solid transparent;transition:all .2s;letter-spacing:.2px;
}
.tab.active{color:var(--accent);border-bottom-color:var(--accent)}
.tab-panel{display:none;padding:16px 16px 0}
.tab-panel.active{display:block}

/* ── CARDS ── */
.card{
  background:var(--surface);border:1px solid var(--border);
  border-radius:var(--radius);padding:16px;margin-bottom:12px;
  transition:border-color .2s;
}
.card:focus-within{border-color:var(--border-active)}
.card-label{
  font-size:10px;font-weight:700;color:var(--text-secondary);
  letter-spacing:1px;text-transform:uppercase;margin-bottom:11px;
  display:flex;align-items:center;gap:6px;
}
.card-label::after{content:'';flex:1;height:1px;background:var(--border)}

/* ── 2FA ── */
.tfa-wrap{display:flex;gap:10px;align-items:center}
.tfa-input-wrap{flex:1;position:relative}
.tfa-input{
  width:100%;background:var(--surface2);border:1.5px solid var(--border);
  border-radius:var(--radius-sm);color:var(--text);
  font-family:'JetBrains Mono',monospace;font-size:13px;
  padding:12px 44px 12px 14px;outline:none;letter-spacing:.5px;
  transition:border-color .2s;
}
.tfa-input:focus{border-color:var(--border-active)}
.tfa-input::placeholder{color:var(--text-dim);font-family:'Inter',sans-serif;letter-spacing:0;font-size:12px}
.tfa-paste-icon{
  position:absolute;right:12px;top:50%;transform:translateY(-50%);
  color:var(--text-secondary);cursor:pointer;font-size:15px;
  transition:color .15s;user-select:none;
}
.tfa-paste-icon:active{color:var(--accent)}
.tfa-display{
  background:var(--surface2);border:1.5px solid var(--border);
  border-radius:var(--radius-sm);min-width:82px;
  padding:10px 12px;text-align:center;cursor:pointer;
  transition:border-color .2s,background .2s;flex-shrink:0;
}
.tfa-display:active{background:var(--accent-glow);border-color:var(--border-active)}
.tfa-display .lbl{font-size:9px;color:var(--text-dim);text-transform:uppercase;letter-spacing:.8px;margin-bottom:3px}
.tfa-display .code{
  font-family:'JetBrains Mono',monospace;font-size:20px;font-weight:700;
  color:var(--accent);letter-spacing:3px;line-height:1;
}
.tfa-display .code.empty{font-size:13px;color:var(--text-dim);letter-spacing:0;font-family:'Inter',sans-serif}
.tfa-timer{height:3px;background:var(--border);border-radius:2px;margin-top:10px;overflow:hidden}
.tfa-timer-bar{height:100%;background:linear-gradient(90deg,var(--accent-dark),var(--accent));border-radius:2px;transition:width .5s linear}
.tfa-footer{display:flex;align-items:center;justify-content:space-between;margin-top:8px}
.tfa-toast{font-size:11px;color:var(--success);min-height:15px;font-weight:500}
.tfa-clear-btn{
  font-size:11px;color:var(--text-dim);background:none;border:none;
  cursor:pointer;padding:3px 0;transition:color .15s;
}
.tfa-clear-btn:hover{color:var(--error)}

/* ── CREDENTIALS ── */
.creds-textarea{
  width:100%;background:var(--surface2);border:1.5px solid var(--border);
  border-radius:var(--radius-sm);color:var(--text);font-family:'JetBrains Mono',monospace;
  font-size:12px;padding:12px 14px;resize:none;height:80px;outline:none;
  transition:border-color .2s;line-height:1.7;
}
.creds-textarea:focus{border-color:var(--border-active)}
.creds-textarea::placeholder{color:var(--text-dim);font-family:'Inter',sans-serif;font-size:12px}
.format-hint{
  font-size:10.5px;color:var(--text-secondary);margin-top:8px;line-height:1.6;
  background:rgba(99,157,255,0.05);border-radius:var(--radius-xs);
  padding:8px 10px;border-left:2px solid var(--accent);
}
.format-hint code{
  font-family:'JetBrains Mono',monospace;font-size:10px;
  color:var(--accent);background:rgba(99,157,255,0.1);
  border-radius:4px;padding:1px 5px;
}
.creds-actions{display:flex;gap:8px;margin-top:10px}
.action-btn{
  flex:1;padding:9px 6px;border:1.5px solid var(--border);border-radius:var(--radius-sm);
  background:var(--surface2);color:var(--text-secondary);
  font-family:'Inter',sans-serif;font-size:11.5px;font-weight:600;
  cursor:pointer;transition:all .18s;text-align:center;
}
.action-btn:hover,.action-btn:active{border-color:var(--accent);color:var(--accent);background:var(--accent-glow)}
.creds-status{font-size:11px;color:var(--success);min-height:16px;margin-top:6px;font-weight:500}

/* ── GET CODE BUTTON ── */
.get-btn{
  width:100%;padding:15px;border:none;border-radius:var(--radius);cursor:pointer;
  font-family:'Inter',sans-serif;font-size:15px;font-weight:700;
  background:linear-gradient(135deg,var(--accent-dark) 0%,var(--accent) 100%);
  color:#fff;transition:opacity .15s,transform .1s,box-shadow .15s;
  margin:4px 0 14px;display:flex;align-items:center;justify-content:center;gap:9px;
  box-shadow:0 4px 16px rgba(99,157,255,0.25);letter-spacing:.2px;
}
.get-btn:active{transform:scale(.98);opacity:.9;box-shadow:none}
.get-btn:disabled{opacity:.4;cursor:not-allowed;transform:none;box-shadow:none}
.spinner{width:16px;height:16px;border:2px solid rgba(255,255,255,.3);
  border-top-color:#fff;border-radius:50%;animation:spin .7s linear infinite;display:none}
@keyframes spin{to{transform:rotate(360deg)}}
.cooldown-label{font-size:12.5px;opacity:.75;font-weight:500}

/* ── OTP RESULTS ── */
.results-wrap{margin-top:2px}
.otp-card{
  background:var(--surface);border:1px solid var(--border);
  border-radius:var(--radius);padding:14px 15px;margin-bottom:10px;
  position:relative;overflow:hidden;
}
.otp-card.latest{border-color:rgba(99,157,255,.3)}
.otp-card.latest::before{
  content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,var(--accent-dark),var(--accent));
}
.otp-tag{
  display:inline-flex;align-items:center;gap:4px;
  font-size:9.5px;font-weight:700;letter-spacing:.5px;text-transform:uppercase;
  color:var(--accent);background:var(--accent-glow);
  border-radius:5px;padding:2px 8px;margin-bottom:9px;
}
.otp-subject{
  font-size:11.5px;color:var(--text-secondary);margin-bottom:9px;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
}
.otp-code-row{display:flex;align-items:center;justify-content:space-between;gap:10px}
.otp-code{
  font-family:'JetBrains Mono',monospace;font-size:22px;font-weight:700;
  color:var(--text);letter-spacing:4px;
}
.copy-btn{
  padding:7px 16px;border:1.5px solid var(--border);border-radius:var(--radius-sm);
  background:var(--surface2);color:var(--text-secondary);
  font-size:12px;font-weight:600;cursor:pointer;transition:all .18s;
  white-space:nowrap;flex-shrink:0;
}
.copy-btn:active,.copy-btn:hover{border-color:var(--accent);color:var(--accent);background:var(--accent-glow)}
.otp-meta{font-size:10.5px;color:var(--text-dim);margin-top:8px}
.error-card{
  background:var(--error-bg);border:1px solid rgba(240,96,128,.2);
  border-radius:var(--radius);padding:13px 15px;margin-bottom:10px;
  font-size:12.5px;color:var(--error);display:flex;align-items:flex-start;gap:8px;line-height:1.5;
}
.empty-state{
  text-align:center;padding:28px 16px;color:var(--text-dim);font-size:13px;
}
.empty-icon{font-size:32px;display:block;margin-bottom:8px;opacity:.5}
</style>
</head>
<body>
<div class="header">
  <div class="header-logo">🛡️</div>
  <h1>Prime Bazar</h1>
  <p>2FA Generator &amp; Email OTP Reader</p>
</div>

<div class="tabs">
  <button class="tab active" onclick="switchTab('tfa',this)">🔐 2FA</button>
  <button class="tab" onclick="switchTab('otp',this)">📧 OTP Email</button>
</div>

<!-- ══ TAB: 2FA ══ -->
<div id="tab-tfa" class="tab-panel active">
  <div class="card">
    <div class="card-label">🔑 Secret Key</div>
    <div class="tfa-wrap">
      <div class="tfa-input-wrap">
        <input class="tfa-input" id="tfaKey" type="text"
          placeholder="2FA secret key paste করুন…"
          autocomplete="off" autocorrect="off" spellcheck="false"
          oninput="onTFAInput()"/>
        <span class="tfa-paste-icon" onclick="pasteTFA()" title="Paste">📋</span>
      </div>
      <div class="tfa-display" id="tfaDisplay" onclick="copyTFACode()">
        <div class="lbl">Code</div>
        <div class="code empty" id="tfaCode">——</div>
      </div>
    </div>
    <div class="tfa-timer"><div class="tfa-timer-bar" id="tfaBar" style="width:100%"></div></div>
    <div class="tfa-footer">
      <button class="tfa-clear-btn" onclick="clearTFA()">✕ Clear</button>
      <div class="tfa-toast" id="tfaToast"></div>
    </div>
  </div>
  <div class="empty-state" id="tfaHint">
    <span class="empty-icon">🔐</span>
    2FA secret key paste করুন —<br/>code auto-generate হবে
  </div>
</div>

<!-- ══ TAB: OTP EMAIL ══ -->
<div id="tab-otp" class="tab-panel">
  <div class="card">
    <div class="card-label">📬 Credentials</div>
    <textarea class="creds-textarea" id="inp"
      placeholder="email@outlook.com|password|refresh_token|client_id"
      autocomplete="off" autocorrect="off" spellcheck="false"
      oninput="onCredsChange()" onpaste="setTimeout(onCredsChange,50)"></textarea>
    <p class="format-hint">
      Format: <code>email</code> | <code>password</code> | <code>refresh_token</code> | <code>client_id</code>
    </p>
    <div class="creds-actions">
      <button class="action-btn" onclick="pasteCreds()">📋 Paste</button>
      <button class="action-btn" onclick="copyCreds()">📄 Copy</button>
      <button class="action-btn" onclick="clearCreds()">🗑 Clear</button>
    </div>
    <div class="creds-status" id="emailStatus"></div>
  </div>

  <button class="get-btn" id="fetchBtn" onclick="getOTP()">
    <span class="spinner" id="spinner"></span>
    <span id="btnLabel">⚡ Get OTP Code</span>
  </button>

  <div id="otpResults" class="results-wrap">
    <div class="empty-state">
      <span class="empty-icon">📧</span>
      Credentials দিয়ে<br/>Get OTP Code চাপুন
    </div>
  </div>
</div>

<script>
const tg=window.Telegram&&window.Telegram.WebApp;
if(tg){tg.ready();tg.expand();}

// ══ TAB SWITCH ══
function switchTab(name,el){
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('tab-'+name).classList.add('active');
}

// ══ 2FA TOTP ══
function b32Decode(s){
  const a='ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
  let bits=0,val=0;const out=[];
  for(const c of s.replace(/=+$/,'').toUpperCase()){
    const i=a.indexOf(c);if(i===-1)continue;
    val=(val<<5)|i;bits+=5;
    if(bits>=8){bits-=8;out.push((val>>bits)&0xff);}
  }
  return new Uint8Array(out);
}
async function genTOTP(secret){
  const key=b32Decode(secret.replace(/\\s/g,''));
  if(!key.length)return null;
  const step=Math.floor(Date.now()/1000/30);
  const buf=new ArrayBuffer(8);
  new DataView(buf).setUint32(4,step,false);
  const ck=await crypto.subtle.importKey('raw',key,{name:'HMAC',hash:'SHA-1'},false,['sign']);
  const sig=new Uint8Array(await crypto.subtle.sign('HMAC',ck,buf));
  const off=sig[19]&0xf;
  const code=(((sig[off]&0x7f)<<24)|(sig[off+1]<<16)|(sig[off+2]<<8)|sig[off+3])%1000000;
  return code.toString().padStart(6,'0');
}

let tfaTimer=null,tfaInterval=null,_lastTFASecret='';
function onTFAInput(){
  clearTimeout(tfaTimer);
  const k=document.getElementById('tfaKey').value.trim();
  _lastTFASecret=k;
  if(!k){resetTFA();return;}
  document.getElementById('tfaHint').style.display='none';
  tfaTimer=setTimeout(()=>refreshTFACode(),300);
}
async function refreshTFACode(){
  const k=_lastTFASecret;
  if(!k)return;
  try{
    const code=await genTOTP(k);
    const el=document.getElementById('tfaCode');
    if(!code){el.textContent='——';el.className='code empty';return;}
    el.textContent=code;el.className='code';
    updateTFABar();
  }catch(e){}
}
function updateTFABar(){
  const rem=(30-(Math.floor(Date.now()/1000)%30))/30*100;
  document.getElementById('tfaBar').style.width=rem+'%';
}
// refresh code every second for live timer
setInterval(()=>{
  if(_lastTFASecret)refreshTFACode();
  if(_lastTFASecret)updateTFABar();
},1000);

function resetTFA(){
  const el=document.getElementById('tfaCode');
  el.textContent='——';el.className='code empty';
  document.getElementById('tfaBar').style.width='100%';
  document.getElementById('tfaHint').style.display='block';
}
function clearTFA(){
  document.getElementById('tfaKey').value='';
  _lastTFASecret='';resetTFA();
  document.getElementById('tfaToast').textContent='';
}
async function pasteTFA(){
  try{
    const t=await navigator.clipboard.readText();
    document.getElementById('tfaKey').value=t.trim();
    _lastTFASecret=t.trim();
    document.getElementById('tfaHint').style.display='none';
    await refreshTFACode();
  }catch(e){document.getElementById('tfaKey').focus();}
}
async function copyTFACode(){
  const code=document.getElementById('tfaCode').textContent;
  if(!code||code==='——')return;
  try{
    await navigator.clipboard.writeText(code);
    const t=document.getElementById('tfaToast');
    t.textContent='✅ '+code+' copied!';
    if(tg&&tg.HapticFeedback)tg.HapticFeedback.notificationOccurred('success');
    setTimeout(()=>{t.textContent='';},2500);
  }catch(e){}
}

// ══ CREDENTIALS ══
function onCredsChange(){
  const raw=document.getElementById('inp').value;
  const parts=raw.split('|');
  const email=(parts[0]||'').trim();
  const st=document.getElementById('emailStatus');
  if(/^[^@\\s]+@[^@\\s]+\\.[^@\\s]+$/.test(email)){
    st.textContent='📧 '+email;
    navigator.clipboard.writeText(email).catch(()=>{});
  } else {
    st.textContent='';
  }
}
async function pasteCreds(){
  try{
    const t=await navigator.clipboard.readText();
    document.getElementById('inp').value=t;
    onCredsChange();
  }catch(e){document.getElementById('inp').focus();}
}
function copyCreds(){
  const v=document.getElementById('inp').value;
  if(!v)return;
  navigator.clipboard.writeText(v).then(()=>{
    const st=document.getElementById('emailStatus');
    st.textContent='✅ Copied!';
    setTimeout(()=>onCredsChange(),1500);
  }).catch(()=>{});
}
function clearCreds(){
  document.getElementById('inp').value='';
  document.getElementById('emailStatus').textContent='';
}

// ══ GET OTP ══
let cooldownActive=false;
async function getOTP(){
  if(cooldownActive)return;
  const raw=document.getElementById('inp').value.trim();
  if(!raw){showError('Credentials দিন।');return;}
  const parts=raw.split('|').map(s=>s.trim());
  if(parts.length<4){showError('Format ঠিক নেই: email|password|refresh_token|client_id');return;}
  setLoading(true);
  try{
    const fd=new FormData();fd.append('data',raw);
    const res=await fetch('/get-otp-api',{method:'POST',body:fd});
    const json=await res.json();
    if(json.error)showError(json.error);
    else addResult(json);
  }catch(e){showError('Network error: '+e.message);}
  finally{setLoading(false);startCooldown();}
}
function setLoading(on){
  const btn=document.getElementById('fetchBtn');
  const sp=document.getElementById('spinner');
  btn.disabled=on;
  sp.style.display=on?'block':'none';
  if(on)document.getElementById('btnLabel').textContent='Fetching…';
  else document.getElementById('btnLabel').textContent='⚡ Get OTP Code';
}
function startCooldown(){
  cooldownActive=true;let secs=3;
  const btn=document.getElementById('fetchBtn');
  const lbl=document.getElementById('btnLabel');
  btn.disabled=true;
  lbl.innerHTML='<span class="cooldown-label">⏳ Wait '+secs+'s</span>';
  const iv=setInterval(()=>{
    secs--;
    if(secs<=0){clearInterval(iv);cooldownActive=false;btn.disabled=false;lbl.textContent='⚡ Get OTP Code';}
    else lbl.innerHTML='<span class="cooldown-label">⏳ Wait '+secs+'s</span>';
  },1000);
}

let otpList=[];
function addResult(d){
  otpList.unshift(d);
  if(otpList.length>5)otpList=otpList.slice(0,5);
  renderResults();
}
function showError(msg){
  const wrap=document.getElementById('otpResults');
  const el=document.createElement('div');
  el.className='error-card';
  el.innerHTML='⚠️ <span>'+esc(msg)+'</span>';
  wrap.prepend(el);
  setTimeout(()=>{if(el.parentNode)el.remove();},6000);
}
function renderResults(){
  const wrap=document.getElementById('otpResults');
  wrap.innerHTML='';
  otpList.forEach((d,i)=>{
    const fmt=d.received?new Date(d.received).toLocaleString('en-BD',{timeZone:'Asia/Dhaka',hour12:true}):'—';
    const div=document.createElement('div');
    div.className='otp-card'+(i===0?' latest':'');
    div.innerHTML=(i===0?'<div class="otp-tag">✨ Latest</div>':'')
      +'<div class="otp-subject">📩 '+esc(d.subject||'—')+'</div>'
      +'<div class="otp-code-row">'
      +  '<span class="otp-code">'+esc(d.code||'—')+'</span>'
      +  '<button class="copy-btn" onclick="copyCode(this,\''+esc(d.code||'')+'\')">📋 Copy</button>'
      +'</div>'
      +'<div class="otp-meta">👤 '+esc(d.sender||'—')+' &nbsp;·&nbsp; 🕒 '+esc(fmt)+'</div>';
    wrap.appendChild(div);
  });
}
async function copyCode(btn,code){
  if(!code)return;
  try{
    await navigator.clipboard.writeText(code);
    btn.textContent='✅ Copied';
    btn.style.borderColor='var(--success)';btn.style.color='var(--success)';
    if(tg&&tg.HapticFeedback)tg.HapticFeedback.notificationOccurred('success');
    setTimeout(()=>{btn.textContent='📋 Copy';btn.style.borderColor='';btn.style.color='';},2000);
  }catch(e){btn.textContent='❌ Failed';setTimeout(()=>{btn.textContent='📋 Copy';},2000);}
}
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
</script>
</body>
</html>"""


@flask_app.route("/code-viewer")
@flask_app.route("/api/code-viewer")
def code_viewer():
    return render_template_string(_OTP_HTML)


@flask_app.route("/get-otp-api", methods=["GET", "POST"])
@flask_app.route("/api/get-otp-api", methods=["GET", "POST"])
def get_otp_api():
    raw   = (request.form.get("data") or request.args.get("data") or "").strip()
    parts = [p.strip() for p in raw.split("|")]
    if len(parts) < 4:
        return jsonify({"error": "Invalid format. Use: email|pass|refresh_token|client_id"})
    refresh_token = parts[2]
    client_id     = parts[3]
    if not refresh_token or not client_id:
        return jsonify({"error": "refresh_token and client_id are required."})
    try:
        access_token = get_access_token(client_id, refresh_token)
        result       = fetch_email_code(access_token)
        return jsonify(result)
    except requests.exceptions.HTTPError as e:
        status = e.response.status_code if e.response else 0
        if status in (400, 401):
            return jsonify({"error": "Invalid or expired refresh_token / client_id."})
        return jsonify({"error": f"HTTP {status}: {str(e)}"})
    except requests.exceptions.Timeout:
        return jsonify({"error": "Request timed out. Please try again."})
    except Exception as e:
        return jsonify({"error": str(e)})


def keep_alive():
    try:
        from config import PORT as _PORT
    except ImportError:
        _PORT = int(os.environ.get("PORT", 3000))
    threading.Thread(target=lambda: flask_app.run(host="0.0.0.0", port=_PORT), daemon=True).start()


# ─────────────────────────────────────────────
#  Core Config  (config.py → fallback to env)
# ─────────────────────────────────────────────
try:
    from config import (
        BOT_TOKEN         as _CFG_BOT_TOKEN,
        ADMIN_ID          as _CFG_ADMIN_ID,
        USER_API_ID       as _CFG_USER_API_ID,
        USER_API_HASH     as _CFG_USER_API_HASH,
        USER_SESSION_STRING as _CFG_USER_SESSION,
        APP_DOMAIN        as _CFG_APP_DOMAIN,
    )
    API_TOKEN    = _CFG_BOT_TOKEN
    ADMIN_ID     = _CFG_ADMIN_ID
    USER_API_ID  = _CFG_USER_API_ID
    USER_API_HASH = _CFG_USER_API_HASH
    USER_SESSION  = _CFG_USER_SESSION
    _REPLIT_DOMAIN = _CFG_APP_DOMAIN
except ImportError:
    API_TOKEN     = os.environ.get("BOT_TOKEN", "")
    ADMIN_ID      = int(os.environ.get("ADMIN_ID", "7522357347"))
    USER_API_ID   = os.environ.get("USER_API_ID", "")
    USER_API_HASH = os.environ.get("USER_API_HASH", "")
    USER_SESSION  = os.environ.get("USER_SESSION_STRING", "")
    _REPLIT_DOMAIN = (
        os.environ.get("APP_DOMAIN") or
        (os.environ.get("REPLIT_DOMAINS", "") or os.environ.get("REPLIT_DEV_DOMAIN", "")).split(",")[0]
    ).strip()

if not API_TOKEN or API_TOKEN == "YOUR_BOT_TOKEN_HERE":
    raise ValueError("BOT_TOKEN সেট করা নেই! config.py ফাইলে BOT_TOKEN এর জায়গায় আপনার token বসান।")

bot = telebot.TeleBot(API_TOKEN, parse_mode=None)

# ═══════════════════════════════════════════════
#  SETTINGS  (MongoDB: settings collection)
# ═══════════════════════════════════════════════
_SETTINGS_DEFAULTS = {
    "support_username":           "@owner_of_pam",
    "usd_rate":                   127,
    "welcome_photo_url":          "",
    "welcome_text":               "",
    "maintenance_mode":           False,
    "maintenance_message":        "🔧 বট এখন রক্ষণাবেক্ষণে আছে। শীঘ্রই ফিরে আসছি!\n🔧 Bot is under maintenance. Coming back soon!",
    "force_join_enabled":         True,
    "language_selection_enabled": True,
    "daily_bonus_enabled":        True,
    "daily_bonus_amount":         5,
    "staff_ids":                  [],
    "new_user_alert":             True,
    "faq_enabled":                False,
    "faq_items":                  [],
    "last_weekly_report":         None,
    "force_join_channels": [
        {"username": "@prime_bazar_update", "url": "https://t.me/prime_bazar_update", "name": "Prime Bazar Update"},
        {"username": "@primeaccmarket",     "url": "https://t.me/primeaccmarket",     "name": "Prime Acc Market"},
    ],
    "payment_methods": {
        "Bkash":   "01848288111",
        "Rocket":  "01949037733",
        "Binance": "PayID: 751370600",
    },
    # ── Supplier Bot (Navigation-based) ──────────────────────────
    "supplier_bot_username":      "",
    "supplier_initial_cmd":       "/start",
    "supplier_vpn_button":        "🛡️ Buy VPN",
    "supplier_duration_buttons":  {
        "03d": "3 Days VPN",
        "07d": "7 Days VPN",
        "15d": "15 Days VPN",
        "30d": "30 Days VPN",
    },
    "supplier_confirm_button":    "✅ Confirm",
    "supplier_stock_cmd":         "/stock",
    "supplier_buy_cmd":           "/buy {product} {qty}",
    # button_map: {my_name: supplier_button_text} for both durations and products
    "button_map":                 {},
    # Master switch — when False, ALL vpn purchases go to manual admin delivery
    "userbot_enabled":            True,
    # VPN display-name list that always requires manual admin delivery
    # (case-insensitive match against the product's display name)
    "manual_delivery_vpns": [
        "14 Days Proton", "30 Days Nord", "30 Days Mysterium", "30 Days Express",
    ],
}

def load_settings():
    return db_load_settings(_SETTINGS_DEFAULTS)

def save_settings(data):
    db_save_settings(data)

settings = load_settings()

def cfg(key):
    return settings.get(key, _SETTINGS_DEFAULTS.get(key))

def SUPPORT_USERNAME():    return cfg("support_username")
def FORCE_JOIN_CHANNELS(): return cfg("force_join_channels")
def PAYMENT_METHODS():     return cfg("payment_methods")
def USD_RATE():            return cfg("usd_rate")


# ═══════════════════════════════════════════════
#  TEXTS  (MongoDB: texts collection)
# ═══════════════════════════════════════════════
_TEXT_FALLBACK = {
    "en": {"welcome": "👋 Welcome to Prime Bazar!", "banned": "🚫 You are banned.", "back": "🔙 Back"},
    "bn": {"welcome": "👋 Prime Bazar-এ স্বাগতম!", "banned": "🚫 ব্যান করা হয়েছে।",  "back": "🔙 পেছনে"},
}

def load_texts():
    return db_load_texts(_TEXT_FALLBACK)

STRINGS = load_texts()

def t(user_id, key, **kwargs):
    uid  = str(user_id)
    lang = user_data.get(uid, {}).get("language", "bn") or "bn"
    text = STRINGS.get(lang, STRINGS["bn"]).get(key) or STRINGS["bn"].get(key, key)
    if kwargs:
        try: text = text.format(**kwargs)
        except Exception: pass
    return text

def s(lang, key, **kwargs):
    text = STRINGS.get(lang, STRINGS["bn"]).get(key) or STRINGS["bn"].get(key, key)
    if kwargs:
        try: text = text.format(**kwargs)
        except Exception: pass
    return text


# ═══════════════════════════════════════════════
#  USED TRX IDs  (MongoDB: used_trxids)
# ═══════════════════════════════════════════════
def load_trxids() -> set:
    return db_load_trxids()

def save_trxids(data: set):
    db_save_trxids(data)

used_trxids: set = load_trxids()


# ═══════════════════════════════════════════════
#  PENDING & REJECTED DEPOSITS
# ═══════════════════════════════════════════════
_PENDING_DEPS_FILE  = "pending_deposits.json"
_REJECTED_DEPS_FILE = "rejected_deposits.json"

def _load_pending_deps() -> dict:
    return db_load_pending_deps()

def _save_pending_deps(data: dict):
    db_save_pending_deps(data)

def _load_rejected_deps() -> list:
    return db_load_rejected_deps()

def _save_rejected_deps(data: list):
    db_save_rejected_deps(data)

pending_deposits: dict  = _load_pending_deps()   # key=uid, val={uid,method,bdt,trx,time}
rejected_deposits: list = _load_rejected_deps()  # [{uid,method,bdt,trx,time}]

_MANUAL_ORDERS_FILE = "pending_manual_orders.json"

def _load_manual_orders() -> dict:
    return db_load_manual_orders()

def _save_manual_orders(data: dict):
    db_save_manual_orders(data)


# ─── JSON fallback file paths ───────────────────
COUPONS_FILE  = "coupons.json"
MARKET_FILE   = "market_data.json"
DATA_FILE     = "user_data.json"
SETTINGS_FILE = "settings.json"
TEXTS_FILE    = "texts.json"

# ═══════════════════════════════════════════════
#  COUPONS  (coupons.json)
# ═══════════════════════════════════════════════
def load_coupons():
    return db_load_coupons()

def save_coupons(data):
    db_save_coupons(data)

coupons = load_coupons()


# ═══════════════════════════════════════════════
#  MARKET DATA  (market_data.json)
# ═══════════════════════════════════════════════
_DEFAULT_VPN_DURS = {
    "03d": {"name": "⚡ 03 Days"},
    "07d": {"name": "🔥 07 Days"},
    "15d": {"name": "💫 15 Days"},
    "30d": {"name": "🚀 30 Days"},
}

# VPN duration auto-detection hints
_VPN_DUR_HINTS = {
    "03d": ["3 day", "03 day", " 3d", "03d"],
    "07d": ["7 day", "07 day", " 7d", "07d", "week", "1 week"],
    "15d": ["15 day", "15d"],
    "30d": ["30 day", "30d", "month", "1 month"],
}

def guess_vpn_dur_key(duration_str):
    """Auto-detect VPN duration key from a duration string."""
    d = (duration_str or "").lower()
    for key, hints in _VPN_DUR_HINTS.items():
        for h in hints:
            if h in d:
                return key
    # fallback: try leading numbers
    nums = re.findall(r'\d+', d)
    if nums:
        n = int(nums[0])
        if n <= 3:  return "03d"
        if n <= 7:  return "07d"
        if n <= 15: return "15d"
        return "30d"
    return "30d"

def load_market_data():
    data = db_load_market(_DEFAULT_VPN_DURS)
    # Auto-fix vpn_dur on any VPN product missing it
    for p_name, p_info in data.get("products", {}).items():
        if p_info.get("cat") == "vpn" and not p_info.get("vpn_dur"):
            p_info["vpn_dur"] = guess_vpn_dur_key(p_info.get("duration", ""))
    return data

def save_market_data(data):
    db_save_market(data)

market_data = load_market_data()

def get_products():      return market_data.get("products", {})
def get_categories():    return market_data.get("categories", {})
def get_vpn_durations(): return market_data.get("vpn_durations", {})

def disp_name(p_info: dict, fallback_key: str) -> str:
    """Display name for a product — falls back to its internal dict key."""
    return (p_info or {}).get("name") or fallback_key

def _unique_product_key(name: str) -> str:
    """Generate a unique internal dict key for a product, allowing duplicate
    display names (e.g. same VPN name in different directions/servers)."""
    prods = market_data.get("products", {})
    if name not in prods:
        return name
    i = 2
    while f"{name} #{i}" in prods:
        i += 1
    return f"{name} #{i}"

def _is_manual_delivery_vpn(p_info: dict, fallback_key: str = "") -> bool:
    """True when this VPN purchase must be delivered manually by the admin.

    Priority order:
      1. Per-product toggle  (p_info["manual_delivery"] = True/False)   ← highest
      2. Global userbot switch  (settings["userbot_enabled"] = False)
      3. Global manual list     (settings["manual_delivery_vpns"])
    """
    import re as _re
    # 1. Per-product toggle — explicit True/False overrides everything
    per_prod = p_info.get("manual_delivery")
    if per_prod is not None:
        return bool(per_prod)
    # 2. Global userbot switch
    if cfg("userbot_enabled") is False:
        return True
    # 3. Global name-based manual list
    manual_names = {n.strip().lower() for n in (cfg("manual_delivery_vpns") or [])}
    name     = (p_info.get("name") or fallback_key or "").strip().lower()
    raw_dur  = (p_info.get("duration") or "").strip().lower()
    duration = _re.sub(r'^[^\w\s]+\s*', '', raw_dur).strip()
    return (
        name in manual_names
        or f"{duration} {name}" in manual_names
        or f"{name} {duration}" in manual_names
        or bool(raw_dur and f"{raw_dur} {name}" in manual_names)
    )

def _load_stock_rows(p_name):
    """Return remaining stock rows (mail accounts) for a product from MongoDB.
    One-time migration: if MongoDB has nothing yet but an old local .xlsx
    file exists (pre-MongoDB era / leftover from a previous deploy), import
    it into MongoDB and delete the local copy so it's never re-imported."""
    rows = db_load_stock(p_name)
    if rows:
        return rows
    fp = get_products().get(p_name, {}).get("file", "")
    if fp and os.path.exists(fp):
        try:
            df = pd.read_excel(fp, header=0)
            rows = df.to_dict("records")
            if rows:
                db_save_stock(p_name, rows)
        except Exception:
            rows = []
        try: os.remove(fp)
        except Exception: pass
        return rows
    return []


def get_stock_count(p_name):
    prods = get_products()
    if p_name not in prods: return 0
    p_info = prods[p_name]
    # VPN with supplier bot connected → virtual stock (supplier manages actual qty)
    if p_info.get("cat") == "vpn" and cfg("supplier_bot_username") and _userbot:
        synced = p_info.get("synced_stock")
        return synced if synced is not None else 99
    # Explicit synced stock count
    synced = p_info.get("synced_stock")
    if synced is not None: return synced
    return len(_load_stock_rows(p_name))


# ═══════════════════════════════════════════════
#  USER DATA  (user_data.json)
# ═══════════════════════════════════════════════
_USER_DEFAULTS = {
    "balance":            0,
    "total_deposit":      0,
    "total_orders":       0,
    "orders":             [],
    "deposit_history":    [],
    "last_deposit_time":  None,
    "last_purchase_time": None,
    "last_daily_bonus":   None,
    "language":           "bn",
    "show_photo":         True,
    "joined":             False,
    "banned":             False,
    "username":           "",
    "first_name":         "",
}

def load_data():
    return db_load_all_users(_USER_DEFAULTS)

def save_data(data):
    """Bulk-save all users. Prefer save_one_user() for single-user updates —
    it's much cheaper than rewriting every user document."""
    db_save_all_users(data)

def save_one_user(uid, data):
    db_save_one_user(str(uid), data)

user_data = load_data()

def get_user(user_id):
    uid = str(user_id)
    if uid not in user_data:
        user_data[uid] = dict(_USER_DEFAULTS); user_data[uid]["orders"] = []; save_data(user_data)
    u = user_data[uid]; changed = False
    for k, v in _USER_DEFAULTS.items():
        if k not in u: u[k] = v; changed = True
    if not u.get("language"): u["language"] = "bn"; changed = True
    if changed: save_data(user_data)
    return u

def get_lang(user_id):
    return user_data.get(str(user_id), {}).get("language", "bn") or "bn"

def is_banned(user_id):
    return user_data.get(str(user_id), {}).get("banned", False)

def update_user(user_id, key, amount):
    uid = str(user_id); get_user(uid)
    user_data[uid][key] = user_data[uid].get(key, 0) + amount
    save_data(user_data)

def bst_now():
    return (datetime.utcnow() + timedelta(hours=6)).strftime("%d-%m-%Y | %I:%M %p")


# ═══════════════════════════════════════════════════════════════════════
#  PYROGRAM USERBOT BRIDGE  (Navigation-based inline button clicking)
# ═══════════════════════════════════════════════════════════════════════
_pyro_loop:      asyncio.AbstractEventLoop           = None
_userbot:        "Client | None"                     = None
_pyro_pending:   "dict[str, asyncio.Future]"         = {}    # for stock-sync
_userbot_ready   = threading.Event()
# uid → {d_name, duration, total, qty, time, user_name, username}
# Persisted to disk so orders survive bot restarts
_pending_manual_orders: dict = _load_manual_orders()

# UIDs currently being processed — prevents double-click duplicate orders
_processing_uids: set = set()

# Queue receives (kind, message) from the supplier bot for navigation
_supplier_queue: "asyncio.Queue | None"              = None
_supplier_nav_lock: "asyncio.Lock | None"            = None


class ProductUnavailableError(Exception):
    """Raised when the supplier bot cannot find the requested product."""

class SupplierTimeoutError(Exception):
    """Raised when the supplier bot does not respond within the allowed time."""


# ── Helpers ───────────────────────────────────────────────────────────

def _is_from_supplier(message) -> bool:
    """Return True if message was sent by the configured supplier bot."""
    supplier_raw = cfg("supplier_bot_username").lstrip("@").lower()
    if not supplier_raw:
        return False
    sender_un = (message.from_user.username or "").lower()
    sender_id  = str(message.from_user.id)
    return sender_un == supplier_raw or sender_id == supplier_raw


def _safe_text(msg) -> str:
    """
    Safely extract text/caption from a Pyrogram message.

    Pyrogram represents `.text`/`.caption` with a custom `Str` type that
    re-splits the string by UTF-16 surrogate pairs on *every* slice/index
    operation (Telegram entity offsets are UTF-16-based). Supplier replies
    are emoji-heavy, and slicing that `Str` (e.g. `text[:200]` for logging)
    can land in the middle of a surrogate pair and raise:
        UnicodeDecodeError: 'utf-16-le' codec can't decode bytes...
    Converting to a plain built-in `str` immediately drops the custom
    slicing behaviour, so all later slicing/regex is safe. We also guard
    the attribute access itself in case Pyrogram's own entity parsing hit
    the same bug while building the message.
    """
    for attr in ("text", "caption"):
        try:
            val = getattr(msg, attr, None)
            if val:
                return str(val)  # plain str -> normal (safe) slicing from here on
        except UnicodeDecodeError as e:
            print(f"[Encoding] Failed to decode supplier message .{attr} ({e}); skipping this field.")
            continue
        except Exception:
            continue
    return ""


def _safe_slice(text: str, limit: int) -> str:
    """Truncate a plain str for logging without ever raising (belt-and-braces)."""
    try:
        return str(text)[:limit]
    except Exception:
        return ""


def _clean_text(t: str) -> str:
    """
    Normalize text for robust button matching:
    1. NFC-normalize (fixes emoji codepoint variants)
    2. Remove invisible chars: variation selectors (FE00-FE0F), zero-width chars
    3. Lowercase + strip whitespace
    """
    t = unicodedata.normalize("NFC", t)
    t = re.sub(r"[\uFE00-\uFE0F\u200B-\u200D\uFEFF\u00A0]", "", t)
    return t.lower().strip()


def _words_only(t: str) -> str:
    """Strip ALL non-alphanumeric chars (emoji, symbols, punct) — keeps only words+digits+spaces."""
    return re.sub(r"[^\w\s]", " ", _clean_text(t), flags=re.UNICODE).split()


def _find_button_by_text(message, target_text: str):
    """
    Find the first button matching target_text in BOTH keyboard types:
      • InlineKeyboardMarkup  (inline_keyboard)
      • ReplyKeyboardMarkup   (keyboard)

    Matching is tried in 4 passes (most → least strict):
      1. Exact after NFC + invisible-char strip
      2. Exact on words-only (strips emoji/symbols — handles Unicode variation selectors)
      3. Partial substring after NFC strip
      4. Partial on words-only

    Returns (row_idx, col_idx, btn_text_str, kb_type) or None.
      kb_type = 'inline' | 'reply'
    """
    if not message or not getattr(message, "reply_markup", None):
        return None
    rm = message.reply_markup

    tgt_clean = _clean_text(target_text)
    tgt_words = _words_only(target_text)

    def _scan(rows, kb_type):
        buttons = [
            (r, c, getattr(btn, "text", None) or str(btn))
            for r, row in enumerate(rows)
            for c, btn in enumerate(row)
        ]
        # Pass 1: exact NFC match
        for r, c, raw in buttons:
            if _clean_text(raw) == tgt_clean:
                return r, c, raw, kb_type
        # Pass 2: exact words-only match (ignores emoji encoding differences)
        for r, c, raw in buttons:
            if _words_only(raw) == tgt_words and tgt_words:
                return r, c, raw, kb_type
        # Pass 3: partial NFC match
        for r, c, raw in buttons:
            bc = _clean_text(raw)
            if tgt_clean in bc or bc in tgt_clean:
                return r, c, raw, kb_type
        # Pass 4: partial words-only match
        for r, c, raw in buttons:
            bw = _words_only(raw)
            if tgt_words and (all(w in bw for w in tgt_words) or all(w in tgt_words for w in bw)):
                return r, c, raw, kb_type
        return None

    # Inline keyboard takes priority (more specific)
    if hasattr(rm, "inline_keyboard") and rm.inline_keyboard:
        res = _scan(rm.inline_keyboard, "inline")
        if res:
            return res

    # Reply keyboard (bottom buttons, e.g. after /start)
    if hasattr(rm, "keyboard") and rm.keyboard:
        res = _scan(rm.keyboard, "reply")
        if res:
            return res

    return None


async def _nav_wait_response(timeout: float = 5, require_keyboard: bool = True):
    """
    Wait for the next message (new or edited) from the supplier bot.
    If require_keyboard=True: prefers messages with InlineKeyboard or ReplyKeyboard,
    but also accepts plain text that looks like credentials or a menu heading.
    If require_keyboard=False: accepts any message from supplier bot.
    """
    if _supplier_queue is None:
        return None
    loop = asyncio.get_running_loop()
    deadline = loop.time() + timeout
    last_plain = None  # remember last plain-text message as fallback
    while True:
        remaining = deadline - loop.time()
        if remaining <= 0:
            return last_plain  # return last plain msg if we have one, else None
        try:
            _kind, msg = await asyncio.wait_for(_supplier_queue.get(), timeout=remaining)
            rm = getattr(msg, "reply_markup", None)
            has_inline = bool(rm and getattr(rm, "inline_keyboard", None))
            has_reply  = bool(rm and getattr(rm, "keyboard", None))
            if has_inline or has_reply:
                return msg  # keyboard message always wins
            if not require_keyboard:
                return msg  # accept any message
            # Plain text — return only if it looks like credentials
            text = _safe_text(msg).lower()
            if any(k in text for k in ("mail", "email", "pass", "pwd", "account", "@", ":")):
                return msg
            # Save as fallback but keep draining for a keyboard message
            last_plain = msg
        except asyncio.TimeoutError:
            return last_plain


async def _click_button(client, message, btn_text: str, timeout: float = 5):
    """
    Locate a button (InlineKeyboard OR ReplyKeyboard) and 'click' it:
      • InlineKeyboard → request_callback_answer()
      • ReplyKeyboard  → send_message(btn_text)   [bottom keyboard buttons]
    Returns (next_message, status_str).
    Status: 'ok' | 'not_found' | 'timeout'
    """
    result = _find_button_by_text(message, btn_text)
    if result is None:
        # Log all available buttons so admin alert can show them
        rm = getattr(message, "reply_markup", None)
        print(f"[NAV] Button '{btn_text}' not found. Available:")
        for rows, label in [
            (getattr(rm, "inline_keyboard", []), "inline"),
            (getattr(rm, "keyboard", []),        "reply"),
        ]:
            for row in (rows or []):
                for b in row:
                    print(f"  [{label}] • '{getattr(b,'text',b)}'")
        return None, "not_found"

    _r, _c, actual_text, kb_type = result
    actual_text = str(actual_text)  # plain str -> avoid Pyrogram's surrogate-slicing Str type
    supplier = cfg("supplier_bot_username")

    try:
        if kb_type == "reply":
            # Reply Keyboard: send the button label as a plain text message
            await client.send_message(supplier, actual_text)
            print(f"[NAV] Sent reply-kb text: '{actual_text}'")
        else:
            # Inline Keyboard: use raw MTProto GetBotCallbackAnswer — most reliable
            rm = message.reply_markup
            btn = rm.inline_keyboard[_r][_c]
            cb_data = (btn.callback_data or "").encode("utf-8")
            try:
                peer = await client.resolve_peer(message.chat.id)
                await client.invoke(
                    raw_fn.messages.GetBotCallbackAnswer(
                        peer=peer,
                        msg_id=message.id,
                        data=cb_data,
                        game=False,
                    )
                )
                print(f"[NAV] Clicked inline-kb (raw): '{actual_text}'")
            except Exception as cb_err:
                print(f"[NAV] Raw callback invoke: '{actual_text}' — {cb_err}")
    except Exception as e:
        print(f"[NAV] click_button error ({kb_type}): {e}")

    await asyncio.sleep(2.0)
    next_msg = await _nav_wait_response(timeout=timeout)
    if next_msg is None:
        return None, "timeout"
    return next_msg, "ok"


# ── Sophisticated Navigation ──────────────────────────────────────────

async def _vpn_buy_nav_async(product_name: str, dur_key: str, qty: int):
    """
    Full navigation-based VPN purchase via supplier bot inline buttons.

    Flow per unit:
      1. Send initial command  →  wait for menu       (5s)
      2. Click VPN category button  →  wait           (5s)
      3. Click duration button  →  wait               (5s)
      4. Click product button  →  wait for confirm    (5s)
      5. Click confirm button  →  wait for delivery   (5s)
      6. Extract Email / Password from delivery message

    Returns:
      list[(email, password)]  on success
      "unavailable"             if product button not found
      "timeout"                 if any step times out
      None                      on connection failure
    """
    if not _userbot or not _supplier_queue or not _supplier_nav_lock:
        return None

    supplier      = cfg("supplier_bot_username")
    initial_cmd   = cfg("supplier_initial_cmd")        or "/start"
    vpn_btn_txt   = cfg("supplier_vpn_button")         or "🛡️ Buy VPN"
    dur_btns      = cfg("supplier_duration_buttons")   or {}
    confirm_txt   = cfg("supplier_confirm_button")     or "✅ Confirm"
    btn_map       = cfg("button_map")                  or {}
    # Resolve duration: use mapped name if set, else fall back to dur_btns entry
    raw_dur_text  = dur_btns.get(dur_key, "")
    dur_btn_text  = btn_map.get(raw_dur_text, raw_dur_text)

    credentials = []

    async with _supplier_nav_lock:
        for _attempt in range(qty):

            # ── Drain stale queue messages before each attempt ─────
            while not _supplier_queue.empty():
                try: _supplier_queue.get_nowait()
                except Exception: break

            # ── Step 1: Send initial command ──────────────────────
            try:
                await _userbot.send_message(supplier, initial_cmd)
            except Exception as e:
                print(f"[NAV] send_message error: {e}")
                return "timeout"

            await asyncio.sleep(2.0)   # give supplier bot time to process /start
            menu_msg = await _nav_wait_response(timeout=10)
            if not menu_msg:
                print("[NAV] ✗ No menu from supplier bot after /start.")
                return "timeout"
            print(f"[NAV] ✓ Menu received: {bool(getattr(menu_msg,'reply_markup',None))}")

            def _nav_alert(step: str, wanted: str, msg_obj=None):
                """
                Send admin a Telegram alert listing all available buttons.
                Collects both InlineKeyboard AND ReplyKeyboard buttons.
                Runs synchronously — called via run_in_executor so it never blocks the loop.
                """
                rm = getattr(msg_obj, "reply_markup", None) if msg_obj else None
                found = []
                if rm:
                    for row in (getattr(rm, "inline_keyboard", None) or []):
                        for b in row:
                            found.append(f"[inline] {getattr(b, 'text', str(b))}")
                    for row in (getattr(rm, "keyboard", None) or []):
                        for b in row:
                            found.append(f"[reply]  {getattr(b, 'text', str(b))}")
                found_txt = "\n".join(f"  • `{t}`" for t in found) or "  _(no buttons)_"
                try:
                    bot.send_message(
                        ADMIN_ID,
                        f"⚠️ *Userbot Nav Failed — Step: {step}*\n"
                        f"✨━━━━━━━━━━━━✨\n"
                        f"🔍 Looking for: `{wanted}`\n"
                        f"📋 *Available buttons:*\n{found_txt}\n"
                        f"✨━━━━━━━━━━━━✨\n"
                        "_Update button name in Admin → Supplier Settings → Duration Buttons or Map Buttons._",
                        parse_mode="Markdown"
                    )
                except Exception as _ae:
                    print(f"[NAV] _nav_alert send error: {_ae}")

            loop = asyncio.get_running_loop()

            # ── Step 2: Click the VPN category button ─────────────
            dur_msg, status = await _click_button(_userbot, menu_msg, vpn_btn_txt, timeout=12)
            if status != "ok":
                print(f"[NAV] ✗ VPN button '{vpn_btn_txt}' → {status}")
                await loop.run_in_executor(None, _nav_alert, "VPN Category", vpn_btn_txt, menu_msg)
                return "unavailable" if status == "not_found" else "timeout"
            print(f"[NAV] ✓ VPN button clicked")

            # ── Step 3: Click the duration button ─────────────────
            # button_map may override the exact text; dur_btn_text is already resolved above.
            # If dur_btn_text is empty (supplier collapses duration into VPN step), skip.
            if dur_btn_text and _find_button_by_text(dur_msg, dur_btn_text):
                prod_list_msg, status = await _click_button(
                    _userbot, dur_msg, dur_btn_text, timeout=12)
                if status != "ok":
                    print(f"[NAV] ✗ Duration button '{dur_btn_text}' → {status}")
                    await loop.run_in_executor(None, _nav_alert, "Duration", dur_btn_text, dur_msg)
                    return "timeout"
                print(f"[NAV] ✓ Duration button clicked")
            else:
                # Duration button not present — supplier may show products directly after VPN btn
                print(f"[NAV] ⚡ Duration step skipped (button not found in current menu)")
                prod_list_msg = dur_msg

            # ── Step 4: Find & click the product button ───────────
            short_name  = product_name.replace(" VPN", "").replace(" vpn", "").strip()
            # Resolve via button_map: full name first, then short name, then raw
            mapped_name = btn_map.get(product_name, btn_map.get(short_name, ""))
            # Build candidate list: mapped override first, then full name, then short name
            candidates  = []
            if mapped_name and mapped_name not in (product_name, short_name):
                candidates.append(mapped_name)
            candidates += [product_name, short_name]
            # De-duplicate while preserving order
            seen = set(); candidates = [c for c in candidates if c and not (c in seen or seen.add(c))]

            prod_found  = False
            confirm_msg = None
            status      = "not_found"

            for candidate in candidates:
                if _find_button_by_text(prod_list_msg, candidate):
                    confirm_msg, status = await _click_button(
                        _userbot, prod_list_msg, candidate, timeout=12)
                    prod_found = True
                    print(f"[NAV] ✓ Product matched via candidate: '{candidate}'")
                    break

            if not prod_found:
                print(f"[NAV] ✗ Product '{product_name}' not in supplier menu. "
                      f"Candidates tried: {candidates}")
                await loop.run_in_executor(
                    None, _nav_alert, "Product", " / ".join(candidates), prod_list_msg)
                return "unavailable"
            if status != "ok" or not confirm_msg:
                print(f"[NAV] ✗ Product click → {status}")
                return "timeout"
            print(f"[NAV] ✓ Product button clicked")

            # ── Step 5: Click the Confirm button ──────────────────
            delivery_msg, status = await _click_button(
                _userbot, confirm_msg, confirm_txt, timeout=15)
            if status != "ok" or not delivery_msg:
                print(f"[NAV] ✗ Confirm button → {status}")
                await loop.run_in_executor(None, _nav_alert, "Confirm", confirm_txt, confirm_msg)
                return "timeout"
            print(f"[NAV] ✓ Confirm clicked — parsing delivery")

            # ── Step 6: Extract credentials ────────────────────────
            raw = _safe_text(delivery_msg)
            print(f"[NAV] Delivery raw text: {_safe_slice(raw, 200)}")

            def _extract_creds(text: str):
                """Try to find (email, password) in a block of text. Returns (mail_m, pass_m)."""
                mail = re.search(
                    r"(?:mail|email|user(?:name)?|id)\s*[:\-|➤→✉️📧]*\s*([a-zA-Z0-9_.+\-]+@[^\s\n,|]+)",
                    text, re.IGNORECASE)
                pw = re.search(
                    r"(?:pass(?:word)?|passwd|pwd|key|🔑)\s*[:\-|➤→🔑]*\s*([^\s\n,|]{4,})",
                    text, re.IGNORECASE)
                if not mail:
                    mail = re.search(
                        r"([a-zA-Z0-9_.+\-]+@[a-zA-Z0-9\-]+\.[a-zA-Z]{2,})", text)
                if not pw:
                    for line in text.splitlines():
                        if re.search(r"pass|pwd|key|🔑", line, re.IGNORECASE):
                            tok = re.findall(r"\S{4,}", line)
                            if tok:
                                pw = tok[-1]
                                break
                    else:
                        pw = None
                else:
                    pw = pw.group(1)
                return (mail.group(1) if mail else None), pw

            mail, pw = _extract_creds(raw)

            # The supplier sometimes confirms the order first ("Order completed")
            # and sends the actual Mail/Pass in a *second*, separate message a
            # moment later. If credentials are still missing, keep listening on
            # the supplier queue briefly for a follow-up message and merge it in.
            if not mail or not pw:
                extra_wait_deadline = asyncio.get_running_loop().time() + 8.0
                while (not mail or not pw) and asyncio.get_running_loop().time() < extra_wait_deadline:
                    remaining = extra_wait_deadline - asyncio.get_running_loop().time()
                    if remaining <= 0:
                        break
                    try:
                        _kind, follow_msg = await asyncio.wait_for(_supplier_queue.get(), timeout=remaining)
                    except asyncio.TimeoutError:
                        break
                    follow_text = _safe_text(follow_msg)
                    if not follow_text:
                        continue
                    print(f"[NAV] Follow-up supplier message: {_safe_slice(follow_text, 200)}")
                    f_mail, f_pw = _extract_creds(follow_text)
                    mail = mail or f_mail
                    pw   = pw or f_pw
                    raw  = raw + "\n" + follow_text

            mail = mail or "—"
            pw   = pw or "—"
            print(f"[NAV] ✓ Parsed — mail={mail}  pass={'***' if pw != '—' else '—'}")
            credentials.append((mail, pw))

    return credentials if credentials else "timeout"


# ── Stock Sync (legacy text-command approach) ─────────────────────────

async def _userbot_send_wait(command: str, timeout: float = 35) -> "str | None":
    """Send a text command to supplier bot and await its first reply (stock sync)."""
    if not _userbot:
        return None
    supplier = cfg("supplier_bot_username")
    if not supplier:
        return None
    req_id = f"{time.monotonic()}"
    fut: asyncio.Future = _pyro_loop.create_future()
    _pyro_pending[req_id] = fut
    try:
        await _userbot.send_message(supplier, command)
        result = await asyncio.wait_for(asyncio.shield(fut), timeout=timeout)
        return str(result) if result else result  # force plain str (safe slicing downstream)
    except asyncio.TimeoutError:
        _pyro_pending.pop(req_id, None)
        return None
    except UnicodeDecodeError as e:
        # Pyrogram's UTF-16 surrogate-pair handling can choke on emoji-heavy
        # supplier replies. Don't let that kill the whole request — log and
        # treat it the same as "no usable reply".
        _pyro_pending.pop(req_id, None)
        print(f"_userbot_send_wait encoding error (ignored): {e}")
        return None
    except Exception as e:
        _pyro_pending.pop(req_id, None)
        print(f"_userbot_send_wait error: {e}")
        return None


async def _check_supplier_balance_async() -> "float | None":
    """
    Two-step balance check:
      1. Send /start  → supplier sends main menu (with ✨ Balance button)
      2. Send '✨ Balance' → supplier replies with account info containing balance
      3. Parse the balance amount from the reply (label-based, or number + currency unit)
    """
    if not _userbot or not cfg("supplier_bot_username"):
        return None
    try:
        # Step 1: /start — just to wake up the supplier bot / reset its state
        await _userbot_send_wait(cfg("supplier_initial_cmd") or "/start", timeout=15)
        await asyncio.sleep(1)   # small pause before clicking the button
        # Step 2: send the Balance button text
        response = await _userbot_send_wait("✨ Balance", timeout=20)
        if not response:
            return None
        response = str(response)  # plain str -> safe slicing, no UTF-16 surrogate bug
        print(f"[BalCheck] raw reply: {_safe_slice(response, 200)}")
        # Parse e.g. "💎✨ Balance: 10 BDT" or "Balance: 250.5 BDT"
        m = re.search(
            r"balance\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*(?:BDT|TK|Tk|Taka|৳)?",
            response, re.IGNORECASE)
        if not m:
            # Fallback: a bare number followed directly by a currency unit,
            # e.g. "10 BDT" / "250.5 Tk" without the word "Balance" at all.
            m = re.search(r"(\d+(?:\.\d+)?)\s*(?:BDT|TK|Tk|৳)", response, re.IGNORECASE)
        if m:
            return float(m.group(1))
    except UnicodeDecodeError as e:
        print(f"[BalCheck] encoding error (ignored): {e}")
    except Exception as e:
        print(f"[BalCheck] error: {e}")
    return None


_LOW_BAL_THRESHOLD = 50   # BDT — alert admin when supplier balance drops below this


async def _sync_stock_async():
    """Ask supplier bot for stock info and parse the text reply."""
    cmd      = cfg("supplier_stock_cmd") or "/stock"
    response = await _userbot_send_wait(cmd)
    if not response:
        return None, None
    response = str(response)  # plain str -> safe slicing downstream
    updates: dict = {}
    for name, count in re.findall(
            r"([^\n:|\-]+)[:\-|]\s*(\d+)\s*(?:pcs|pc|pieces|left|available)?",
            response, re.IGNORECASE):
        name = name.strip()
        if name:
            updates[name] = int(count)
    return updates, response


# ── Pyrogram Startup ──────────────────────────────────────────────────

async def _pyro_main():
    global _userbot, _supplier_queue, _supplier_nav_lock
    if not PYRO_OK:
        print("⚠️  Userbot disabled — Pyrogram not installed.")
        _userbot_ready.set(); return
    if not USER_API_ID or not USER_API_HASH:
        print("⚠️  Userbot disabled — USER_API_ID / USER_API_HASH not set.")
        _userbot_ready.set(); return
    if not USER_SESSION:
        print("⚠️  Userbot disabled — USER_SESSION_STRING not set.")
        _userbot_ready.set(); return

    _supplier_queue    = asyncio.Queue()
    _supplier_nav_lock = asyncio.Lock()

    try:
        _userbot = Client(
            name           = "userbot",
            api_id         = int(USER_API_ID),
            api_hash       = USER_API_HASH,
            session_string = USER_SESSION,
            no_updates     = False,
        )

        @_userbot.on_message(pf.private)
        async def _on_msg(client, message):
            if not _is_from_supplier(message):
                return
            # Debug: always log the raw supplier message so encoding/parsing
            # issues are visible in the logs (safe — never raises).
            print(f"Supplier Raw Message: {_safe_slice(_safe_text(message), 500)}")
            # Feed navigation queue
            if _supplier_queue is not None:
                await _supplier_queue.put(("new", message))
            # Also resolve legacy text-command futures (stock sync)
            for req_id, fut in list(_pyro_pending.items()):
                if not fut.done():
                    fut.set_result(_safe_text(message))
                    _pyro_pending.pop(req_id, None)
                    break

        @_userbot.on_edited_message(pf.private)
        async def _on_edit(client, message):
            if not _is_from_supplier(message):
                return
            if _supplier_queue is not None:
                await _supplier_queue.put(("edit", message))

        await _userbot.start()
        me   = await _userbot.get_me()
        name = me.first_name or "Userbot"
        un   = f"@{me.username}" if me.username else "(no username)"
        print(f"✅ Userbot connected: {un} ({name})")
        _userbot_ready.set()
        while True:
            await asyncio.sleep(60)

    except Exception as exc:
        print(f"⚠️  Userbot error: {exc}")
        _userbot = None
        _userbot_ready.set()


def _start_pyro_thread():
    """Dedicated asyncio event loop for Pyrogram (run in background thread)."""
    global _pyro_loop
    _pyro_loop = asyncio.new_event_loop()
    asyncio.set_event_loop(_pyro_loop)
    _pyro_loop.create_task(_pyro_main())
    _pyro_loop.run_forever()


def _run_async(coro, timeout=40):
    """Run a Pyrogram coroutine from a Telebot (sync) thread."""
    if _pyro_loop is None or not _pyro_loop.is_running():
        return None
    fut = asyncio.run_coroutine_threadsafe(coro, _pyro_loop)
    try:
        return fut.result(timeout=timeout)
    except Exception as e:
        print(f"_run_async error: {e}")
        return None


def userbot_status() -> str:
    if not PYRO_OK:                 return "❌ Pyrogram not installed"
    if not USER_API_ID:             return "❌ USER_API_ID not set"
    if _userbot is None:
        if not _userbot_ready.is_set(): return "⏳ Connecting..."
        return "❌ Not connected"
    return "✅ Connected"


# ═══════════════════════════════════════════════
#  FORCE JOIN
# ═══════════════════════════════════════════════
def check_membership(user_id):
    if not cfg("force_join_enabled"): return True
    channels = FORCE_JOIN_CHANNELS()
    if not channels: return True
    for ch in channels:
        try:
            m = bot.get_chat_member(ch["username"], user_id)
            if m.status not in ("member", "administrator", "creator"):
                return False
        except Exception:
            return False
    return True

def send_join_prompt(chat_id, lang="bn"):
    mk = types.InlineKeyboardMarkup(row_width=1)
    for ch in FORCE_JOIN_CHANNELS():
        mk.add(types.InlineKeyboardButton(f"📢 {ch['name']}", url=ch["url"]))
    mk.add(types.InlineKeyboardButton(
        STRINGS.get(lang, STRINGS["bn"]).get("joined_btn", "✅ I've Joined!"),
        callback_data="check_join"))
    bot.send_message(chat_id,
        STRINGS.get(lang, STRINGS["bn"]).get("join_required", "🔐 Please join our channels first!"),
        reply_markup=mk, parse_mode="Markdown")


# ─────────────────────────────────────────────
#  Guard
# ─────────────────────────────────────────────
def guard(message):
    uid  = str(message.chat.id)
    lang = get_lang(uid)
    if is_banned(uid):
        bot.send_message(message.chat.id,
            STRINGS.get(lang, STRINGS["bn"]).get("banned", "🚫 You are banned."),
            parse_mode="Markdown"); return False
    if cfg("maintenance_mode") and message.chat.id != ADMIN_ID:
        bot.send_message(message.chat.id, cfg("maintenance_message")); return False
    if not check_membership(message.chat.id):
        send_join_prompt(message.chat.id, lang); return False
    # Auto-save username and first_name for admin search
    try:
        fu = message.from_user
        if fu:
            u = get_user(uid)
            changed = False
            new_un = f"@{fu.username}" if fu.username else ""
            new_fn = fu.first_name or ""
            if new_un and u.get("username") != new_un:
                user_data[uid]["username"] = new_un; changed = True
            if new_fn and u.get("first_name") != new_fn:
                user_data[uid]["first_name"] = new_fn; changed = True
            if changed: save_data(user_data)
    except Exception:
        pass
    return True


# ─────────────────────────────────────────────
#  Main Menu
# ─────────────────────────────────────────────
def main_menu(user_id):
    lang = get_lang(user_id)
    S    = STRINGS.get(lang, STRINGS["bn"])
    # Admin-configurable labels (settings.json) take priority over language strings
    lbl_profile = cfg("btn_profile") or S.get("btn_profile", "🟣 👤 Profile")
    lbl_deposit = cfg("btn_deposit") or S.get("btn_deposit", "🟣 💰 Deposit")
    lbl_shop    = cfg("btn_shop")    or S.get("btn_shop",    "🟣 🛒 Shop Now")
    lbl_price   = cfg("btn_price")   or S.get("btn_price",   "🟣 💎 Price List")
    lbl_support = cfg("btn_support") or S.get("btn_support", "🟢 ☎️ Support")
    lbl_daily   = cfg("btn_daily")   or S.get("btn_daily",   "🎁 Daily Bonus")
    mk = types.ReplyKeyboardMarkup(row_width=2, resize_keyboard=True)
    mk.add(types.KeyboardButton(lbl_profile),
           types.KeyboardButton(lbl_deposit))
    mk.add(types.KeyboardButton(lbl_shop),
           types.KeyboardButton(lbl_price))
    if cfg("daily_bonus_enabled"):
        mk.add(types.KeyboardButton(lbl_daily),
               types.KeyboardButton(lbl_support))
    else:
        mk.add(types.KeyboardButton(lbl_support))
    return mk

def _btn_labels(*keys):
    """Return all known text values for the given button keys.
    Includes both language-string variants AND the admin-configured label from settings.json."""
    labels = set()
    for ls in STRINGS.values():
        for k in keys:
            v = ls.get(k)
            if v: labels.add(v)
    # Also include admin-overridden labels so message routing still works after a rename
    for k in keys:
        v = cfg(k)
        if v: labels.add(v)
    return labels


# ═══════════════════════════════════════════════
#  /start  /menu
# ═══════════════════════════════════════════════
# ═══════════════════════════════════════════════════════════
#  /get_code — Hotmail/Outlook OTP Reader (Loop Mode)
# ═══════════════════════════════════════════════════════════

_otp_cred_cache: dict = {}   # uid → {email, password, refresh_token, client_id}


def _otp_bst_time(utc_str: str) -> str:
    """ISO UTC string → Bangladesh Standard Time (UTC+6)."""
    try:
        from datetime import timezone, timedelta as _td
        dt = datetime.strptime(utc_str[:19], "%Y-%m-%dT%H:%M:%S")
        dt = dt.replace(tzinfo=timezone.utc) + _td(hours=6)
        return dt.strftime("%d %b %Y, %I:%M %p BST")
    except Exception:
        return utc_str


def _otp_fetch_and_reply(chat_id: int, uid: str) -> bool:
    """
    Fetch OTP using cached credentials and send result to chat_id.
    Returns True if OTP was found, False otherwise.
    """
    creds = _otp_cred_cache.get(uid)
    if not creds:
        bot.send_message(chat_id,
            "❌ Credentials পাওয়া গেলো না। /get\\_code দিয়ে আবার শুরু করুন।",
            parse_mode="Markdown")
        return False

    # "Checking…" spinner
    spin = bot.send_message(
        chat_id,
        "✨━━━━━━━━━━━━━━✨\n"
        "⏳ *Inbox চেক করা হচ্ছে...*\n"
        "✨━━━━━━━━━━━━━━✨",
        parse_mode="Markdown"
    )

    try:
        access_token = get_access_token(creds["client_id"], creds["refresh_token"])
        result       = fetch_email_code(access_token)
    except requests.exceptions.HTTPError as e:
        try: bot.delete_message(chat_id, spin.message_id)
        except Exception: pass
        status = e.response.status_code if e.response else 0
        if status in (400, 401):
            err_msg = "❌ *Token Expired / Invalid*\nRefresh token মেয়াদোত্তীর্ণ বা Client ID ভুল। নতুন credentials দিন।"
        else:
            err_msg = f"❌ *HTTP Error {status}:* `{str(e)[:150]}`"
        mk = types.InlineKeyboardMarkup()
        mk.add(types.InlineKeyboardButton("🔄 Retry", callback_data=f"otp_refresh|{uid}"))
        bot.send_message(chat_id,
            f"✨━━━━━━━━━━━━━━✨\n{err_msg}\n✨━━━━━━━━━━━━━━✨",
            parse_mode="Markdown", reply_markup=mk)
        return False
    except requests.exceptions.Timeout:
        try: bot.delete_message(chat_id, spin.message_id)
        except Exception: pass
        mk = types.InlineKeyboardMarkup()
        mk.add(types.InlineKeyboardButton("🔄 Retry", callback_data=f"otp_refresh|{uid}"))
        bot.send_message(chat_id,
            "✨━━━━━━━━━━━━━━✨\n"
            "❌ *Network Timeout*\nInternet সমস্যা। একটু পরে Retry করুন।\n"
            "✨━━━━━━━━━━━━━━✨",
            parse_mode="Markdown", reply_markup=mk)
        return False
    except Exception as e:
        try: bot.delete_message(chat_id, spin.message_id)
        except Exception: pass
        mk = types.InlineKeyboardMarkup()
        mk.add(types.InlineKeyboardButton("🔄 Retry", callback_data=f"otp_refresh|{uid}"))
        bot.send_message(chat_id,
            f"✨━━━━━━━━━━━━━━✨\n"
            f"❌ *Error:* `{str(e)[:200]}`\n"
            f"✨━━━━━━━━━━━━━━✨",
            parse_mode="Markdown", reply_markup=mk)
        return False

    try: bot.delete_message(chat_id, spin.message_id)
    except Exception: pass

    if "error" in result:
        # OTP not found — show error + Refresh button
        mk = types.InlineKeyboardMarkup()
        mk.add(types.InlineKeyboardButton("🔄 Refresh", callback_data=f"otp_refresh|{uid}"))
        bot.send_message(chat_id,
            f"✨━━━━━━━━━━━━━━✨\n"
            f"📭 *কোনো OTP পাওয়া যায়নি*\n"
            f"✨━━━━━━━━━━━━━━✨\n"
            f"📧 *Email:* `{creds['email']}`\n\n"
            f"⚠️ _{result['error']}_\n\n"
            f"🔄 *Refresh* বাটনে ক্লিক করুন অথবা নতুন credentials পাঠান।",
            parse_mode="Markdown", reply_markup=mk)
        return False
    else:
        # ✅ OTP found — premium result UI
        bst = _otp_bst_time(result.get("received", ""))
        bot.send_message(chat_id,
            f"✨━━━━━━━━━━━━━━✨\n"
            f"🔐 *OTP Code পাওয়া গেছে!*\n"
            f"✨━━━━━━━━━━━━━━✨\n"
            f"📧 *Email:* `{creds['email']}`\n"
            f"📩 *Subject:* {result.get('subject', '—')}\n"
            f"👤 *From:* `{result.get('sender', '—')}`\n"
            f"🕒 *Received:* {bst}\n"
            f"✨━━━━━━━━━━━━━━✨\n\n"
            f"🔑 *আপনার OTP Code:*\n\n"
            f"`{result['code']}`\n\n"
            f"✨━━━━━━━━━━━━━━✨\n"
            f"_📋 Code-এ ট্যাপ করলে কপি হবে_",
            parse_mode="Markdown")
        # Clear cache — next credentials will be fresh
        _otp_cred_cache.pop(uid, None)
        return True


def _otp_send_loop_prompt(chat_id: int):
    """Send 'paste next credentials' prompt (loop mode)."""
    bot.send_message(
        chat_id,
        "✨━━━━━━━━━━━━━━✨\n"
        "☎️ *Loop Mode — পরবর্তী Email*\n"
        "✨━━━━━━━━━━━━━━✨\n"
        "📋 পরবর্তী credentials পাঠান:\n\n"
        "`email|password|refresh_token|client_id`\n\n"
        "🔴 বন্ধ করতে /start বা /menu পাঠান।",
        parse_mode="Markdown"
    )


@bot.message_handler(commands=["get_code"])
def cmd_get_code(message):
    uid = str(message.chat.id)
    _otp_cred_cache.pop(uid, None)   # clear any stale cache
    bot.send_message(
        message.chat.id,
        "✨━━━━━━━━━━━━━━✨\n"
        "☎️ *Hotmail / Outlook OTP Reader*\n"
        "✨━━━━━━━━━━━━━━✨\n"
        "📋 নিচের format-এ credentials পাঠান:\n\n"
        "`email|password|refresh_token|client_id`\n\n"
        "📌 *উদাহরণ:*\n"
        "`user@outlook.com|Pass123|0.AXoA...|abc-def`\n\n"
        "🔴 বন্ধ করতে /start বা /menu পাঠান।",
        parse_mode="Markdown"
    )
    bot.register_next_step_handler(message, _otp_handle_creds)


def _otp_handle_creds(message):
    """Process credentials pasted by the user, fetch OTP, then loop."""
    uid = str(message.chat.id)
    txt = (message.text or "").strip()

    # Exit loop on commands
    if txt.startswith("/start") or txt.startswith("/menu"):
        welcome(message)
        return

    # Validate format
    parts = txt.split("|")
    if len(parts) < 4:
        bot.send_message(
            message.chat.id,
            "✨━━━━━━━━━━━━━━✨\n"
            "⚠️ *ভুল Format!*\n"
            "✨━━━━━━━━━━━━━━✨\n"
            "📋 সঠিক format:\n"
            "`email|password|refresh_token|client_id`\n\n"
            "_আবার পাঠান অথবা /start দিয়ে মেনুতে যান।_",
            parse_mode="Markdown"
        )
        bot.register_next_step_handler(message, _otp_handle_creds)
        return

    email_val   = parts[0].strip()
    password    = parts[1].strip()
    ref_token   = parts[2].strip()
    client_id   = "|".join(p.strip() for p in parts[3:])  # client_id may contain |

    _otp_cred_cache[uid] = {
        "email":         email_val,
        "password":      password,
        "refresh_token": ref_token,
        "client_id":     client_id,
    }

    found = _otp_fetch_and_reply(message.chat.id, uid)

    # Always loop: show prompt and wait for next credentials
    _otp_send_loop_prompt(message.chat.id)
    bot.register_next_step_handler(message, _otp_handle_creds)


@bot.callback_query_handler(func=lambda c: c.data.startswith("otp_refresh|"))
def otp_refresh_cb(call):
    """🔄 Refresh button — re-check same credentials."""
    uid = call.data.split("|", 1)[1]
    bot.answer_callback_query(call.id, "🔄 Re-checking inbox...")
    # Delete the old error card
    try: bot.delete_message(call.message.chat.id, call.message.message_id)
    except Exception: pass

    _otp_fetch_and_reply(call.message.chat.id, uid)

    # Re-enter loop
    _otp_send_loop_prompt(call.message.chat.id)
    bot.register_next_step_handler(call.message, _otp_handle_creds)


# ─────────────────────────────────────────────────────────────

@bot.message_handler(commands=["start", "menu"])
def welcome(message):
    uid  = str(message.chat.id)
    user = get_user(uid)
    if is_banned(uid):
        bot.send_message(message.chat.id,
            STRINGS.get(get_lang(uid), STRINGS["bn"]).get("banned", "🚫 Banned."),
            parse_mode="Markdown"); return
    if cfg("maintenance_mode") and message.chat.id != ADMIN_ID:
        bot.send_message(message.chat.id, cfg("maintenance_message")); return
    if not check_membership(message.chat.id):
        send_join_prompt(message.chat.id, get_lang(uid)); return
    _is_new = not user_data.get(uid, {}).get("joined", False)
    user_data[uid]["joined"] = True; save_data(user_data)
    if _is_new and cfg("new_user_alert") is not False:
        try:
            fn = message.from_user.first_name or ""
            un = f"@{message.from_user.username}" if message.from_user.username else "—"
            bot.send_message(ADMIN_ID,
                f"👋 *নতুন ইউজার যোগ দিয়েছে!*\n"
                f"✨━━━━━━━━━━━━━━━━━━✨\n"
                f"👤 নাম: *{fn}*\n"
                f"🔖 Username: {un}\n"
                f"🆔 ID: `{uid}`\n"
                f"✨━━━━━━━━━━━━━━━━━━✨",
                parse_mode="Markdown")
        except Exception:
            pass
    lang = user.get("language")
    if not lang: _send_lang_selection(message.chat.id); return
    _send_welcome(message.chat.id, uid, lang)


def _send_welcome(chat_id, uid, lang):
    welcome_txt = cfg("welcome_text") or STRINGS.get(lang, STRINGS["bn"]).get("welcome", "👋 Welcome!")
    photo_url   = cfg("welcome_photo_url")
    if photo_url:
        try:
            bot.send_photo(chat_id, photo_url, caption=welcome_txt,
                           reply_markup=main_menu(uid), parse_mode="Markdown"); return
        except Exception: pass
    bot.send_message(chat_id, welcome_txt, reply_markup=main_menu(uid), parse_mode="Markdown")


def _send_lang_selection(chat_id):
    if not cfg("language_selection_enabled"):
        uid = str(chat_id)
        user_data[uid]["language"] = "bn"; save_data(user_data)
        _send_welcome(chat_id, uid, "bn"); return
    mk = types.InlineKeyboardMarkup(row_width=2)
    mk.add(types.InlineKeyboardButton("🇬🇧 English", callback_data="setlang|en"),
           types.InlineKeyboardButton("🇧🇩 বাংলা",   callback_data="setlang|bn"))
    bot.send_message(chat_id,
        STRINGS["bn"].get("select_lang", "🌐 Select your language:"),
        reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data == "check_join")
def verify_join(call):
    uid  = str(call.message.chat.id)
    lang = get_lang(uid)
    if not check_membership(call.message.chat.id):
        bot.answer_callback_query(call.id,
            STRINGS.get(lang, STRINGS["bn"]).get("not_joined", "❌ Not joined yet!"),
            show_alert=True); return
    bot.answer_callback_query(call.id, "✅")
    user_data[uid]["joined"] = True; save_data(user_data)
    try: bot.delete_message(call.message.chat.id, call.message.message_id)
    except Exception: pass
    user = get_user(uid)
    if not user.get("language"): _send_lang_selection(call.message.chat.id)
    else: _send_welcome(call.message.chat.id, uid, user["language"])


@bot.callback_query_handler(func=lambda c: c.data.startswith("setlang|"))
def set_language(call):
    uid  = str(call.message.chat.id)
    lang = call.data.split("|")[1]
    get_user(uid); user_data[uid]["language"] = lang; save_data(user_data)
    bot.answer_callback_query(call.id, "✅")
    try: bot.delete_message(call.message.chat.id, call.message.message_id)
    except Exception: pass
    bot.send_message(call.message.chat.id,
        STRINGS[lang].get("lang_saved", "✅ Language set!"), parse_mode="Markdown")
    _send_welcome(call.message.chat.id, uid, lang)


# ═══════════════════════════════════════════════
#  PROFILE
# ═══════════════════════════════════════════════
@bot.message_handler(func=lambda m: m.text in _btn_labels("btn_profile"))
def profile(message):
    if not guard(message): return
    uid  = str(message.chat.id)
    u    = get_user(uid)
    lang = get_lang(uid)
    S    = STRINGS.get(lang, STRINGS["bn"])

    first     = message.from_user.first_name or ""
    last      = message.from_user.last_name  or ""
    full_name = f"{first} {last}".strip() or "Unknown"
    username  = f"@{message.from_user.username}" if message.from_user.username else "—"
    last_dep  = u.get("last_deposit_time")  or S.get("no_transaction", "N/A")
    last_buy  = u.get("last_purchase_time") or S.get("no_transaction", "N/A")

    caption = (
        f"╔══════════════════════╗\n"
        f"      {S.get('profile_title', '👤 USER PROFILE')}\n"
        f"╚══════════════════════╝\n\n"
        f"{S.get('field_name',   '✨ *Name:*')} {full_name}\n"
        f"{S.get('field_user',   '🔖 *Username:*')} {username}\n"
        f"{S.get('field_id',     '🆔 *User ID:*')} `{uid}`\n\n"
        f"💠━━━━━━━━━━━━━━━━━━💠\n"
        f"{S.get('field_bal',    '💰 *Balance:*')} *{u['balance']:.2f} BDT*\n"
        f"{S.get('field_dep',    '📈 *Total Deposit:*')} *{u['total_deposit']:.2f} BDT*\n"
        f"{S.get('field_orders', '🛒 *Total Orders:*')} *{u['total_orders']}*\n"
        f"💠━━━━━━━━━━━━━━━━━━💠\n"
        f"{S.get('field_lastdep','⏳ *Last Deposit:*')}\n   🕐 {last_dep}\n"
        f"{S.get('field_lastbuy','🎁 *Last Purchase:*')}\n   🕐 {last_buy}\n"
        f"✨━━━━━━━━━━━━━━━━━━✨"
    )

    edit_lbl   = "✏️ Edit Profile" if lang == "en" else "✏️ প্রোফাইল এডিট"
    photo_lbl  = ("🖼 Hide Photo" if u.get("show_photo", True) else "🖼 Show Photo") if lang == "en" \
                 else ("🖼 ফটো লুকান" if u.get("show_photo", True) else "🖼 ফটো দেখান")
    hist_lbl   = "📋 Order History" if lang == "en" else "📋 অর্ডার হিস্টোরি"
    resend_lbl = "🔄 Resend Last Order" if lang == "en" else "🔄 শেষ অর্ডার পুনরায়"
    mk = types.InlineKeyboardMarkup(row_width=2)
    mk.add(types.InlineKeyboardButton(edit_lbl,    callback_data="edit_profile"),
           types.InlineKeyboardButton(photo_lbl,   callback_data="toggle_photo"))
    mk.add(types.InlineKeyboardButton(hist_lbl,    callback_data="order_history"),
           types.InlineKeyboardButton(resend_lbl,  callback_data="resend_last_order"))
    dep_hist_lbl = "💳 Deposit History" if lang == "en" else "💳 ডিপোজিট হিস্টোরি"
    mk.add(types.InlineKeyboardButton(dep_hist_lbl, callback_data="deposit_history"))

    if not u.get("show_photo", True):
        bot.send_message(message.chat.id, caption, parse_mode="Markdown", reply_markup=mk); return

    try:
        avatar_img = None
        photos = bot.get_user_profile_photos(message.chat.id, limit=1)
        if photos.total_count > 0:
            file_id  = photos.photos[0][-1].file_id
            file_url = f"https://api.telegram.org/file/bot{bot.token}/{bot.get_file(file_id).file_path}"
            avatar_img = Image.open(io.BytesIO(requests.get(file_url, timeout=10).content))
        if avatar_img is None:
            initials   = ((first[:1] + last[:1]).upper()) or "PB"
            avatar_img = Image.open(io.BytesIO(requests.get(
                f"https://ui-avatars.com/api/?name={initials}&size=256"
                f"&background=1a1a2e&color=e94560&bold=true&format=png", timeout=10).content))
        card_buf = _make_profile_card(avatar_img, full_name, username, uid,
                                      u["balance"], u["total_deposit"], u["total_orders"],
                                      last_dep, last_buy)
        bot.send_photo(message.chat.id, card_buf, caption=caption,
                       parse_mode="Markdown", reply_markup=mk)
    except Exception:
        bot.send_message(message.chat.id, caption, parse_mode="Markdown", reply_markup=mk)


def _make_profile_card(avatar_img, full_name, username, uid,
                       balance, total_deposit, total_orders, last_dep, last_buy):
    W, H   = 520, 580
    ACCENT = (233, 69, 96); GOLD = (255, 200, 80); TEXT_MAIN = (235, 235, 255)
    TEXT_DIM = (140, 140, 180); DIVIDER = (60, 60, 100)
    card   = Image.new("RGB", (W, H), (10, 10, 28))
    d      = ImageDraw.Draw(card)
    try:
        bold_sm = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 15)
        reg     = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 13)
        sm      = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 12)
    except Exception:
        bold_sm = reg = sm = ImageFont.load_default()
    # Avatar circle
    ring_sz = 90; ax = (W - ring_sz*2)//2; ay = 30
    d.ellipse([ax-4, ay-4, ax+ring_sz*2+4, ay+ring_sz*2+4], fill=ACCENT)
    d.ellipse([ax, ay, ax+ring_sz*2, ay+ring_sz*2], fill=(10,10,28))
    try:
        av = avatar_img.resize((ring_sz*2, ring_sz*2)).convert("RGBA")
        mask = Image.new("L", (ring_sz*2, ring_sz*2), 0)
        ImageDraw.Draw(mask).ellipse([0,0,ring_sz*2-1,ring_sz*2-1], fill=255)
        card.paste(av, (ax, ay), mask)
    except Exception: pass
    y  = ay + ring_sz*2 + 12
    safe_name = "".join(c for c in full_name if ord(c) < 0x0600 or c == " ") or full_name[:22]
    nb = d.textbbox((0,0), safe_name[:26], font=bold_sm)
    d.text(((W-(nb[2]-nb[0]))//2, y), safe_name[:26], font=bold_sm, fill=TEXT_MAIN); y += 22
    ub = d.textbbox((0,0), username[:28], font=sm)
    d.text(((W-(ub[2]-ub[0]))//2, y), username[:28], font=sm, fill=ACCENT); y += 20
    d.line([(40, y), (W-40, y)], fill=DIVIDER, width=1); y += 12
    rows = [("ID", uid), ("Balance", f"{balance} BDT"),
            ("Total Deposit", f"{total_deposit} BDT"), ("Total Orders", str(total_orders)),
            ("Last Deposit", str(last_dep)[:22]), ("Last Purchase", str(last_buy)[:22])]
    for i, (label, value) in enumerate(rows):
        d.rectangle([28, y-2, W-28, y+22], fill=(26,26,54) if i%2==0 else (22,22,48))
        d.text((44, y), label, font=bold_sm, fill=TEXT_DIM)
        vb = d.textbbox((0,0), value, font=reg)
        d.text((W-44-(vb[2]-vb[0]), y), value, font=reg, fill=TEXT_MAIN); y += 26
    d.line([(40, y+4), (W-40, y+4)], fill=ACCENT, width=1)
    buf = io.BytesIO(); card.save(buf, format="PNG"); buf.seek(0)
    return buf


@bot.callback_query_handler(func=lambda c: c.data == "edit_profile")
def edit_profile_menu(call):
    uid  = str(call.message.chat.id)
    lang = get_lang(uid)
    txt  = ("✏️ *Edit Profile*\n💎━━━━━━━━━━━━━━━💎\n🌐 Switch language:"
            if lang == "en" else
            "✏️ *প্রোফাইল এডিট*\n💎━━━━━━━━━━━━━━━💎\n🌐 ভাষা পরিবর্তন:")
    mk   = types.InlineKeyboardMarkup(row_width=2)
    mk.add(types.InlineKeyboardButton("🇬🇧 English", callback_data="setlang|en"),
           types.InlineKeyboardButton("🇧🇩 বাংলা",   callback_data="setlang|bn"))
    bot.answer_callback_query(call.id)
    bot.send_message(call.message.chat.id, txt, parse_mode="Markdown", reply_markup=mk)


@bot.callback_query_handler(func=lambda c: c.data == "toggle_photo")
def toggle_photo(call):
    uid = str(call.message.chat.id)
    u   = get_user(uid)
    u["show_photo"] = not u.get("show_photo", True); save_data(user_data)
    lang  = get_lang(uid)
    state = ("shown ✅" if u["show_photo"] else "hidden 🙈") if lang == "en" \
            else ("দেখানো হবে ✅" if u["show_photo"] else "লুকানো হবে 🙈")
    bot.answer_callback_query(call.id, f"Profile photo {state}")


# ═══════════════════════════════════════════════
#  PRICE LIST
# ═══════════════════════════════════════════════
def _duration_sort_key(dur_str: str):
    """Parse a free-text duration ('1 Month', '7 Days', '12-36 Month',
    'Lifetime') into a sortable (bucket, days) tuple so shorter durations
    always come before longer ones, with 'Lifetime'/unlimited last."""
    if not dur_str:
        return (0, 0.0)
    s = dur_str.strip().lower()
    if "lifetime" in s or "unlimited" in s:
        return (2, 0.0)
    nums = re.findall(r"\d+\.?\d*", s)
    n = float(nums[0]) if nums else 0.0
    if "year" in s:
        mult = 365
    elif "month" in s:
        mult = 30
    elif "week" in s:
        mult = 7
    elif "hour" in s:
        mult = 1 / 24
    elif "day" in s:
        mult = 1
    else:
        mult = 30
    return (1, n * mult)


@bot.message_handler(func=lambda m: m.text in _btn_labels("btn_price"))
def price_list(message):
    if not guard(message): return
    uid  = str(message.chat.id)
    lang = get_lang(uid)
    S    = STRINGS.get(lang, STRINGS["bn"])
    cats = get_categories(); prods = get_products()
    txt  = S.get("price_title", "💎 *PRICE LIST*") + "\n✨━━━━━━━━━━━━━━━━━━✨\n"
    for cat_key, cat_info in cats.items():
        cat_items = [(p_name, p_info) for p_name, p_info in prods.items()
                     if p_info.get("cat") == cat_key]
        if not cat_items:
            continue
        # Shortest duration first, then cheapest price first for equal durations
        cat_items.sort(key=lambda item: (
            _duration_sort_key(item[1].get("duration", "")),
            item[1].get("price", 0),
        ))
        txt += f"\n{cat_info.get('emoji','')} *{cat_info['name']}*\n"
        txt += "┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈\n"
        for idx, (p_name, p_info) in enumerate(cat_items, 1):
            dur = f" _{p_info['duration']}_" if p_info.get("duration") else ""
            txt += f"  `{idx:>2}.` *{p_name}*{dur} — 💵 *{p_info.get('price',0)} BDT*\n"
    txt += f"\n✨━━━━━━━━━━━━━━━━━━✨\n{S.get('price_rate','💵 *Binance Rate:* $1 = {{rate}} BDT').format(rate=USD_RATE())}"
    bot.send_message(message.chat.id, txt, parse_mode="Markdown")


# ═══════════════════════════════════════════════
#  SUPPORT
# ═══════════════════════════════════════════════
@bot.message_handler(func=lambda m: m.text in _btn_labels("btn_support"))
def support(message):
    uid  = str(message.chat.id)
    lang = get_lang(uid)
    S    = STRINGS.get(lang, STRINGS["bn"])
    su   = SUPPORT_USERNAME()
    mk   = types.InlineKeyboardMarkup()
    mk.add(types.InlineKeyboardButton("💬 Contact Support", url=f"https://t.me/{su.lstrip('@')}"))
    bot.send_message(message.chat.id,
        S.get("support_text", "☎️ *Support*\n👤 {username}").format(username=su),
        reply_markup=mk, parse_mode="Markdown")


# ═══════════════════════════════════════════════
#  SHOP
# ═══════════════════════════════════════════════
@bot.message_handler(func=lambda m: m.text in _btn_labels("btn_shop"))
def shop_now(message):
    if not guard(message): return
    uid  = str(message.chat.id)
    S    = STRINGS.get(get_lang(uid), STRINGS["bn"])
    cats = get_categories()
    mk   = types.InlineKeyboardMarkup(row_width=2)
    btns = [types.InlineKeyboardButton(cat_info["name"], callback_data=f"cat|{cat_key}")
            for cat_key, cat_info in cats.items()]
    if btns: mk.add(*btns)
    bot.send_message(message.chat.id,
        S.get("shop_cat","🛒 *Select a Category:*"), reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("cat|"))
def show_items(call):
    uid  = str(call.message.chat.id)
    S    = STRINGS.get(get_lang(uid), STRINGS["bn"])
    cat  = call.data.split("|")[1]
    if cat == "vpn":
        _show_vpn_durations(call); return
    if cat == "mail":
        _show_mail_products(call); return
    prods = get_products()
    mk    = types.InlineKeyboardMarkup(row_width=2)
    btns  = [types.InlineKeyboardButton(
                 f"{disp_name(p_info, p_name)} 📦 {get_stock_count(p_name)}pcs", callback_data=f"buy|{p_name}")
             for p_name, p_info in prods.items() if p_info.get("cat") == cat]
    if btns: mk.add(*btns)
    mk.add(types.InlineKeyboardButton(S.get("back","🔙 Back"), callback_data="back_to_cat"))
    bot.edit_message_text(S.get("shop_product","🛍️ *Select a Product:*"),
        call.message.chat.id, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    bot.answer_callback_query(call.id)


# ═══════════════════════════════════════════════
#  MAIL FLOW  (Retail text delivery / Bulk xlsx)
# ═══════════════════════════════════════════════

def _show_mail_products(call):
    uid   = str(call.message.chat.id)
    S     = STRINGS.get(get_lang(uid), STRINGS["bn"])
    prods = get_products()
    mk    = types.InlineKeyboardMarkup(row_width=2)
    btns  = [types.InlineKeyboardButton(
                 f"{p_name} 📦 {get_stock_count(p_name)}pcs",
                 callback_data=f"mailtype|{p_name}")
             for p_name, p_info in prods.items() if p_info.get("cat") == "mail"]
    if btns: mk.add(*btns)
    mk.add(types.InlineKeyboardButton(S.get("back","🔙 Back"), callback_data="back_to_cat"))
    bot.edit_message_text(
        S.get("shop_product","🛍️ *Select a Product:*"),
        call.message.chat.id, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    bot.answer_callback_query(call.id)


@bot.callback_query_handler(func=lambda c: c.data.startswith("mailtype|"))
def mail_type_select(call):
    uid    = str(call.message.chat.id)
    p_name = call.data.split("|", 1)[1]
    prods  = get_products()
    stock  = get_stock_count(p_name)
    lang   = get_lang(uid)
    S      = STRINGS.get(lang, STRINGS["bn"])
    if stock == 0:
        bot.answer_callback_query(call.id, S.get("shop_no_stock","❌ Out of Stock!"), show_alert=True); return
    price = prods[p_name]["price"]
    txt = (
        f"📧 *{p_name}*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"💸 *Price:* {price} BDT/each\n"
        f"📦 *Stock:* {stock} pcs\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"🛒 Retail (1–6) — delivered as *text*\n"
        f"📦 Bulk (any qty) — delivered as *xlsx file*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"Select purchase type:"
    )
    mk = types.InlineKeyboardMarkup(row_width=2)
    mk.add(
        types.InlineKeyboardButton("🛒 Retail Mail", callback_data=f"mailretail|{p_name}"),
        types.InlineKeyboardButton("📦 Bulk Mail",   callback_data=f"mailbulk|{p_name}"),
    )
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="cat|mail"))
    bot.edit_message_text(txt, call.message.chat.id, call.message.message_id,
                          reply_markup=mk, parse_mode="Markdown")
    bot.answer_callback_query(call.id)


@bot.callback_query_handler(func=lambda c: c.data.startswith("mailretail|"))
def mail_retail_qty(call):
    uid    = str(call.message.chat.id)
    p_name = call.data.split("|", 1)[1]
    prods  = get_products()
    stock  = get_stock_count(p_name)
    price  = prods[p_name]["price"]
    max_q  = min(stock, 6)
    txt = (
        f"🛒 *Retail Mail — {p_name}*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"💸 *Price:* {price} BDT/each\n"
        f"📦 *Stock:* {stock} pcs\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"Select quantity (1–{max_q}):"
    )
    mk = types.InlineKeyboardMarkup(row_width=3)
    btns = [types.InlineKeyboardButton(str(q), callback_data=f"mailorder|{p_name}|{q}|retail")
            for q in range(1, max_q + 1)]
    if btns: mk.add(*btns)
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data=f"mailtype|{p_name}"))
    bot.edit_message_text(txt, call.message.chat.id, call.message.message_id,
                          reply_markup=mk, parse_mode="Markdown")
    bot.answer_callback_query(call.id)


@bot.callback_query_handler(func=lambda c: c.data.startswith("mailbulk|"))
def mail_bulk_ask_qty(call):
    uid    = str(call.message.chat.id)
    p_name = call.data.split("|", 1)[1]
    prods  = get_products()
    stock  = get_stock_count(p_name)
    price  = prods[p_name]["price"]
    bot.answer_callback_query(call.id)
    txt = (
        f"📦 *Bulk Mail — {p_name}*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"💸 *Price:* {price} BDT/each\n"
        f"📦 *Stock:* {stock} pcs\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"🔢 Enter quantity (1–{stock}):"
    )
    msg = bot.send_message(call.message.chat.id, txt, parse_mode="Markdown")
    bot.register_next_step_handler(msg, _mail_bulk_qty_input, p_name, stock)


def _mail_bulk_qty_input(message, p_name, stock):
    uid  = str(message.chat.id)
    lang = get_lang(uid)
    S    = STRINGS.get(lang, STRINGS["bn"])
    try:
        qty = int(message.text.strip())
    except (ValueError, AttributeError):
        bot.send_message(uid, S.get("shop_invalid_qty","❌ Please enter a valid number.")); return
    if qty < 1 or qty > stock:
        bot.send_message(uid,
            S.get("shop_qty_err","❌ Quantity must be between 1 and {stock}.").format(stock=stock)); return
    prods = get_products()
    price = prods[p_name]["price"]
    total = round(qty * price, 2)
    u     = get_user(uid)
    txt = (
        f"📦 *Bulk Order Summary*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"📧 *Product:* {p_name}\n"
        f"🔢 *Quantity:* {qty}\n"
        f"💸 *Total:* {total} BDT\n"
        f"💰 *Your Balance:* {u['balance']:.2f} BDT\n"
        f"✨━━━━━━━━━━━━━━━━━━✨"
    )
    mk = types.InlineKeyboardMarkup()
    mk.add(types.InlineKeyboardButton(
        S.get("shop_confirm","✅ Confirm Purchase"),
        callback_data=f"mailorder|{p_name}|{qty}|bulk"))
    mk.add(types.InlineKeyboardButton("❌ Cancel", callback_data="back_to_cat"))
    bot.send_message(uid, txt, reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("mailorder|"))
def mail_order_confirm(call):
    """Show confirm screen for retail; process immediately for bulk (already confirmed)."""
    parts  = call.data.split("|")
    p_name = parts[1]; qty = int(parts[2]); mode = parts[3]
    uid    = str(call.message.chat.id)
    lang   = get_lang(uid)
    S      = STRINGS.get(lang, STRINGS["bn"])
    prods  = get_products()
    if p_name not in prods:
        bot.answer_callback_query(call.id, S.get("shop_file_err","❌ File Error!"), show_alert=True); return
    price = prods[p_name]["price"]
    total = round(qty * price, 2)
    u     = get_user(uid)
    if u["balance"] < total:
        bot.answer_callback_query(call.id,
            S.get("shop_low_bal","❌ Insufficient balance!").format(bal=u["balance"], total=total),
            show_alert=True); return

    if mode == "retail":
        txt = (
            f"🛒 *Retail Order Summary*\n"
            f"✨━━━━━━━━━━━━━━━━━━✨\n"
            f"📧 *Product:* {p_name}\n"
            f"🔢 *Quantity:* {qty}\n"
            f"💸 *Total:* {total} BDT\n"
            f"💰 *Your Balance:* {u['balance']:.2f} BDT\n"
            f"✨━━━━━━━━━━━━━━━━━━✨"
        )
        mk = types.InlineKeyboardMarkup()
        mk.add(types.InlineKeyboardButton(
            S.get("shop_confirm","✅ Confirm Purchase"),
            callback_data=f"mailpay|{p_name}|{qty}|retail"))
        mk.add(types.InlineKeyboardButton("🔙 Back", callback_data=f"mailretail|{p_name}"))
        bot.edit_message_text(txt, call.message.chat.id, call.message.message_id,
                              reply_markup=mk, parse_mode="Markdown")
        bot.answer_callback_query(call.id)
    else:
        # Bulk: the qty-input message cannot be edited (it's a send_message), go to mailpay directly
        if uid in _processing_uids:
            bot.answer_callback_query(call.id, "⏳ আপনার অর্ডারটি ইতিমধ্যে প্রক্রিয়া হচ্ছে, অনুগ্রহ করে অপেক্ষা করুন।", show_alert=True); return
        _processing_uids.add(uid)
        bot.answer_callback_query(call.id, "⏳ Processing...")
        try:
            _mail_execute(call, uid, p_name, qty, total, mode, S)
        finally:
            _processing_uids.discard(uid)


@bot.callback_query_handler(func=lambda c: c.data.startswith("mailpay|"))
def mail_pay(call):
    parts  = call.data.split("|")
    p_name = parts[1]; qty = int(parts[2]); mode = parts[3]
    uid    = str(call.message.chat.id)
    lang   = get_lang(uid)
    S      = STRINGS.get(lang, STRINGS["bn"])
    if uid in _processing_uids:
        bot.answer_callback_query(call.id, "⏳ আপনার অর্ডারটি ইতিমধ্যে প্রক্রিয়া হচ্ছে, অনুগ্রহ করে অপেক্ষা করুন।", show_alert=True); return
    prods  = get_products()
    if p_name not in prods:
        bot.answer_callback_query(call.id, S.get("shop_file_err","❌ File Error!"), show_alert=True); return
    total = round(qty * prods[p_name]["price"], 2)
    u     = get_user(uid)
    if u["balance"] < total:
        bot.answer_callback_query(call.id,
            S.get("shop_low_bal","❌ Insufficient balance!").format(bal=u["balance"], total=total),
            show_alert=True); return
    _processing_uids.add(uid)
    bot.answer_callback_query(call.id, "⏳ Processing...")
    try:
        _mail_execute(call, uid, p_name, qty, total, mode, S)
    finally:
        _processing_uids.discard(uid)


def _mail_execute(call, uid, p_name, qty, total, mode, S):
    """Deduct balance, deliver, update stats, notify admin."""
    try:
        if mode == "retail":
            _deliver_mail_text(uid, p_name, qty, total, S)
        else:
            _deliver_mail_xlsx(uid, p_name, qty, total, S)
        # Deduct & update stats
        update_user(uid, "balance", -total)
        update_user(uid, "total_orders", 1)
        now = bst_now()
        order_id = uuid.uuid4().hex[:8]
        user_data[uid]["orders"].append({"order_id": order_id, "product": p_name, "qty": qty, "total": total, "date": now})
        user_data[uid]["last_purchase_time"] = now
        save_data(user_data)
        _prompt_review(uid, order_id, p_name, S)
        u       = get_user(uid)
        new_bal = round(u["balance"], 2)
        try:
            bot.send_message(ADMIN_ID,
                f"✅ *New Mail Sale!*\n✨━━━━━━━━━━━━✨\n"
                f"👤 User: `{uid}`\n📧 Product: *{p_name}*\n"
                f"🔢 Qty: {qty} | Mode: {mode}\n"
                f"💸 Amount: *{total} BDT*\n"
                f"💰 Remaining Balance: {new_bal} BDT\n🕐 Time: {now}",
                parse_mode="Markdown")
        except Exception: pass
        stock = get_stock_count(p_name)
        if stock < 5:
            try:
                bot.send_message(ADMIN_ID,
                    f"⚠️ *লো স্টক অ্যালার্ট!*\n✨━━━━━━━━━━━━✨\n"
                    f"📧 *{p_name}* — মাত্র *{stock}* পিস বাকি!",
                    parse_mode="Markdown")
            except Exception: pass
        try: bot.delete_message(uid, call.message.message_id)
        except Exception: pass
    except Exception as e:
        print(f"Mail order error: {e}")
        try:
            bot.send_message(ADMIN_ID,
                f"🚨 *Failed Mail Order*\n✨━━━━━━━━━━━━✨\n"
                f"👤 User: `{uid}` | 📧 {p_name} | ❌ {e}",
                parse_mode="Markdown")
        except Exception: pass
        bot.send_message(uid, S.get("shop_file_err","❌ *File Error!* Contact support."),
                         parse_mode="Markdown")


def _low_stock_alert(p_name, remaining):
    """স্টক ২০-এর নিচে নামলে Admin-কে একবার সতর্ক করে।"""
    threshold = 20
    if remaining < threshold:
        try:
            bot.send_message(
                ADMIN_ID,
                f"⚠️ *Low Stock Alert!*\n"
                f"✨━━━━━━━━━━━━━━━━━━✨\n"
                f"📦 প্রোডাক্ট: *{p_name}*\n"
                f"🔢 বাকি স্টক: *{remaining} পিস*\n"
                f"✨━━━━━━━━━━━━━━━━━━✨\n"
                f"📥 দ্রুত স্টক আপলোড করুন!",
                parse_mode="Markdown"
            )
        except Exception:
            pass


def _record_sold_rows(p_name, sold_rows, uid):
    """Tag delivered account rows with buyer/date metadata and archive them
    in the 'sold_stock' collection so the admin can later export used vs.
    fresh mail lists as .xlsx."""
    if not sold_rows:
        return
    now = bst_now()
    u   = get_user(uid)
    tagged = [
        {**row, "Buyer_UID": uid, "Buyer_Name": u.get("name", "?"), "Sold_Date": now}
        for row in sold_rows
    ]
    db_append_sold(p_name, tagged)


def _deliver_mail_text(uid, p_name, qty, total, S):
    """Retail delivery: read Column A, send as text (one account per line).
    Stock is stored in MongoDB, not a local file, so it survives Railway restarts."""
    rows = _load_stock_rows(p_name)
    if not rows:
        raise ValueError(f"Not enough stock: have 0, need {qty}")
    first_col = list(rows[0].keys())[0]
    available = [r.get(first_col) for r in rows if r.get(first_col) is not None]
    if len(available) < qty:
        raise ValueError(f"Not enough stock: have {len(available)}, need {qty}")
    to_send   = available[:qty]
    remaining = rows[qty:]
    db_save_stock(p_name, remaining)
    _low_stock_alert(p_name, len(remaining))
    _record_sold_rows(p_name, rows[:qty], uid)
    accounts_text = "\n".join(f"`{str(a)}`" for a in to_send)
    msg = (
        f"✅ *Purchase Successful!*\n"
        f"💸 Cost: *{total} BDT*\n"
        f"📧 *{p_name}* × {qty}\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"{accounts_text}\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"🎉 Enjoy your account(s)!"
    )
    bot.send_message(uid, msg, parse_mode="Markdown")


def _deliver_mail_xlsx(uid, p_name, qty, total, S):
    """Bulk delivery: send as .xlsx file named [Product_Name]_[Qty]_accounts.xlsx.
    Stock is stored in MongoDB, not a local file, so it survives Railway restarts."""
    rows      = _load_stock_rows(p_name)
    to_send   = rows[:qty]
    remaining = rows[qty:]
    db_save_stock(p_name, remaining)
    _low_stock_alert(p_name, len(remaining))
    _record_sold_rows(p_name, to_send, uid)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame(to_send).to_excel(writer, index=False)
    out.seek(0)
    fname   = f"{p_name.replace(' ', '_')}_{qty}_accounts.xlsx"
    caption = (
        f"✅ *Bulk Purchase Successful!*\n"
        f"💸 Cost: *{total} BDT*\n"
        f"📧 *{p_name}* × {qty}\n"
        f"🎉 Enjoy your accounts!"
    )
    bot.send_document(uid, out, visible_file_name=fname, caption=caption, parse_mode="Markdown")


def _show_vpn_durations(call):
    uid  = str(call.message.chat.id)
    S    = STRINGS.get(get_lang(uid), STRINGS["bn"])
    durs = get_vpn_durations()
    mk   = types.InlineKeyboardMarkup(row_width=2)
    btns = [types.InlineKeyboardButton(di["name"], callback_data=f"vpndur|{dk}")
            for dk, di in durs.items()]
    if btns: mk.add(*btns)
    mk.add(types.InlineKeyboardButton(S.get("back","🔙 Back"), callback_data="back_to_cat"))
    bot.edit_message_text(
        "🛡️ *Buy VPN*\n💎━━━━━━━━━━━━━━━━━━💎\n🕐 Select a duration plan:\n💎━━━━━━━━━━━━━━━━━━💎",
        call.message.chat.id, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    bot.answer_callback_query(call.id)


@bot.callback_query_handler(func=lambda c: c.data.startswith("vpndur|"))
def show_vpn_by_duration(call):
    uid     = str(call.message.chat.id)
    S       = STRINGS.get(get_lang(uid), STRINGS["bn"])
    dur_key = call.data.split("|")[1]
    durs    = get_vpn_durations()
    prods   = get_products()
    dur_name = durs.get(dur_key, {}).get("name", dur_key)
    mk      = types.InlineKeyboardMarkup(row_width=2)
    btns    = [types.InlineKeyboardButton(
                   f"🛡️ {disp_name(p_info, p_name)}", callback_data=f"buy|{p_name}")
               for p_name, p_info in prods.items()
               if p_info.get("cat") == "vpn" and p_info.get("vpn_dur") == dur_key]
    if btns: mk.add(*btns)
    else: mk.add(types.InlineKeyboardButton("❌ No products yet", callback_data="noop"))
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="cat|vpn"))
    bot.edit_message_text(
        f"🛡️ *VPN — {dur_name}*\n💎━━━━━━━━━━━━━━━━━━💎\nChoose a product:\n💎━━━━━━━━━━━━━━━━━━💎",
        call.message.chat.id, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    bot.answer_callback_query(call.id)


@bot.callback_query_handler(func=lambda c: c.data == "back_to_cat")
def back_to_cat(call):
    uid  = str(call.message.chat.id)
    S    = STRINGS.get(get_lang(uid), STRINGS["bn"])
    cats = get_categories()
    mk   = types.InlineKeyboardMarkup(row_width=2)
    btns = [types.InlineKeyboardButton(ci["name"], callback_data=f"cat|{ck}")
            for ck, ci in cats.items()]
    if btns: mk.add(*btns)
    bot.edit_message_text(S.get("shop_cat","🛒 *Select a Category:*"),
        call.message.chat.id, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    bot.answer_callback_query(call.id)


@bot.callback_query_handler(func=lambda c: c.data == "noop")
def noop_cb(call): bot.answer_callback_query(call.id)


@bot.callback_query_handler(func=lambda c: c.data.startswith("buy|"))
def ask_qty(call):
    """Skip quantity prompt — always buy 1. Go straight to confirm page."""
    uid    = str(call.message.chat.id)
    S      = STRINGS.get(get_lang(uid), STRINGS["bn"])
    p_name = call.data.split("|", 1)[1]
    prods  = get_products()
    u      = get_user(uid)
    stock  = get_stock_count(p_name)
    if stock == 0:
        bot.answer_callback_query(call.id, S.get("shop_no_stock","❌ Out of Stock!"), show_alert=True); return
    qty    = 1
    price  = prods[p_name]["price"]
    total  = round(qty * price, 2)
    is_vpn = prods[p_name].get("cat") == "vpn"
    mk     = types.InlineKeyboardMarkup()
    mk.add(types.InlineKeyboardButton(
        S.get("shop_confirm","✅ Confirm Purchase"), callback_data=f"pay|{p_name}|{qty}"))
    if is_vpn:
        pi = prods[p_name]
        d_name = disp_name(pi, p_name)
        duration = pi.get("duration", "N/A")
        order_txt = (
            f"🛡️ *VPN Order Summary*\n✨━━━━━━━━━━━━━━━━━━✨\n"
            f"📦 *Product:* {d_name}\n"
            f"⏳ *Duration:* {duration}\n"
            f"💸 *Price:* *{total} BDT*\n"
            f"💰 *Your Balance:* {u['balance']:.2f} BDT\n✨━━━━━━━━━━━━━━━━━━✨"
        )
    else:
        order_txt = S.get("shop_order","📋 *Order Summary:*\n📦 {product} × {qty}\n💸 {total} BDT").format(
            product=p_name, qty=qty, total=total)
    bot.send_message(call.message.chat.id, order_txt, reply_markup=mk, parse_mode="Markdown")
    bot.answer_callback_query(call.id)


@bot.callback_query_handler(func=lambda c: c.data.startswith("pay|"))
def finalize_order(call):
    parts  = call.data.split("|")
    p_name = parts[1]; qty = int(parts[2])
    uid    = str(call.message.chat.id)
    S      = STRINGS.get(get_lang(uid), STRINGS["bn"])
    if uid in _processing_uids:
        bot.answer_callback_query(call.id, "⏳ আপনার অর্ডারটি ইতিমধ্যে প্রক্রিয়া হচ্ছে, অনুগ্রহ করে অপেক্ষা করুন।", show_alert=True); return
    prods  = get_products()
    if p_name not in prods:
        bot.answer_callback_query(call.id, S.get("shop_file_err","❌ File Error!"), show_alert=True); return
    total = round(qty * prods[p_name]["price"], 2)
    bal   = get_user(uid)["balance"]
    if bal < total:
        bot.answer_callback_query(call.id,
            S.get("shop_low_bal","❌ Insufficient balance! {bal} < {total}").format(bal=bal, total=total),
            show_alert=True); return
    is_vpn = prods[p_name].get("cat") == "vpn"
    _processing_uids.add(uid)
    bot.answer_callback_query(call.id, "⏳ Processing...")
    order_id = uuid.uuid4().hex[:8]
    try:
        p_info = prods[p_name]
        if is_vpn and _is_manual_delivery_vpn(p_info, p_name):
            _deliver_vpn_manual(uid, p_name, p_info, qty, total, S, order_id)
        elif is_vpn:
            d_name   = disp_name(p_info, p_name)
            duration = p_info.get("duration", "N/A")
            success_msg = (
                f"✅ *Purchase Successful!*\n"
                f"✨━━━━━━━━━━━━━━━━━━✨\n"
                f"🛡️ *{d_name}*\n"
                f"⏳ Duration: *{duration}*\n"
                f"💸 Cost: *{total} BDT*\n"
                f"✨━━━━━━━━━━━━━━━━━━✨"
            )
            _deliver_vpn(uid, call, p_name, qty, total, success_msg, S)
        else:
            # Non-VPN non-mail fallback (future categories)
            raise ValueError("Unknown product category — use the mail flow.")
        # Update stats (only reaches here on success)
        update_user(uid, "balance", -total)
        update_user(uid, "total_orders", 1)
        now = bst_now()
        is_manual_vpn = is_vpn and _is_manual_delivery_vpn(p_info, p_name)
        user_data[uid]["orders"].append({"order_id": order_id, "product": p_name, "qty": qty, "total": total, "date": now})
        user_data[uid]["last_purchase_time"] = now
        save_data(user_data)
        if not is_manual_vpn:
            # Manual VPN orders get their review prompt only after the admin
            # actually delivers credentials (see _adm_send_vpn_deliver).
            _prompt_review(uid, order_id, p_name, S)
        # Admin success notification
        u = get_user(uid)
        new_bal = round(u["balance"], 2)
        try:
            bot.send_message(ADMIN_ID,
                f"✅ *New Sale!*\n"
                f"✨━━━━━━━━━━━━✨\n"
                f"👤 User: `{uid}` | {u.get('name','?')}\n"
                f"📦 Product: *{p_name}*\n"
                f"💸 Amount: *{total} BDT*\n"
                f"💰 Remaining Balance: {new_bal} BDT\n"
                f"🕐 Time: {now}",
                parse_mode="Markdown")
        except Exception: pass
        rem_stock = get_stock_count(p_name)
        if rem_stock < 5:
            bot.send_message(ADMIN_ID,
                f"⚠️ *লো স্টক অ্যালার্ট!*\n✨━━━━━━━━━━━━✨\n📦 *{p_name}* — মাত্র *{rem_stock}* পিস বাকি!",
                parse_mode="Markdown")
        try: bot.delete_message(uid, call.message.message_id)
        except Exception: pass
    except ProductUnavailableError:
        u = get_user(uid)
        try:
            bot.send_message(ADMIN_ID,
                f"🚨 *Failed Order — Product Unavailable*\n"
                f"✨━━━━━━━━━━━━✨\n"
                f"👤 User: `{uid}` | {u.get('name','?')}\n"
                f"📦 Product: *{p_name}*\n"
                f"❌ Reason: Supplier has no matching button\n"
                f"💰 Charged: No",
                parse_mode="Markdown")
        except Exception: pass
        bot.send_message(uid,
            "❌ *Product Currently Unavailable!*\n"
            "✨━━━━━━━━━━━━✨\n"
            "The supplier does not have this product right now.\n"
            "💰 *No charge has been made to your balance.*\n"
            "✨━━━━━━━━━━━━✨",
            parse_mode="Markdown")
        try: bot.delete_message(uid, call.message.message_id)
        except Exception: pass
    except SupplierTimeoutError as ste:
        u = get_user(uid)
        try:
            bot.send_message(ADMIN_ID,
                f"🚨 *Failed Order — Supplier Issue*\n"
                f"✨━━━━━━━━━━━━✨\n"
                f"👤 User: `{uid}` | {u.get('name','?')}\n"
                f"📦 Product: *{p_name}*\n"
                f"❌ Reason: {ste}\n"
                f"💰 Charged: No",
                parse_mode="Markdown")
        except Exception: pass
        support_url = cfg("support_username") or "@owner_of_pam"
        support_link = f"https://t.me/{support_url.lstrip('@')}"
        mk_err = types.InlineKeyboardMarkup()
        mk_err.add(types.InlineKeyboardButton("🆘 Support", url=support_link))
        bot.send_message(uid,
            "⚠️ *সাময়িক অসুবিধার জন্য আমরা দুঃখিত।*\n"
            "✨━━━━━━━━━━━━✨\n"
            "আপনার সমস্যাটি নিচে *Support* বাটনে ক্লিক করে সাপোর্টে জানান।\n\n"
            "💰 *আপনার ব্যালেন্স থেকে কোনো টাকা কাটা হয়নি।*\n"
            "✨━━━━━━━━━━━━✨",
            reply_markup=mk_err,
            parse_mode="Markdown")
        try: bot.delete_message(uid, call.message.message_id)
        except Exception: pass
    except Exception as e:
        print(f"Order error: {e}")
        try:
            u = get_user(uid)
            bot.send_message(ADMIN_ID,
                f"🚨 *Failed Order — Unknown Error*\n"
                f"✨━━━━━━━━━━━━✨\n"
                f"👤 User: `{uid}` | {u.get('name','?')}\n"
                f"📦 Product: *{p_name}*\n"
                f"❌ Error: {e}\n"
                f"💰 Charged: No",
                parse_mode="Markdown")
        except Exception: pass
        bot.send_message(uid, S.get("shop_file_err","❌ *File Error!* Contact support."), parse_mode="Markdown")
    finally:
        _processing_uids.discard(uid)


def _deliver_mail(uid, call, p_name, qty, total, success_msg, S):
    """Deliver mail accounts as .xlsx — stock lives in MongoDB, not a local
    file, so it survives Railway restarts/redeploys."""
    rows      = _load_stock_rows(p_name)
    to_send   = rows[:qty]
    remaining = rows[qty:]
    db_save_stock(p_name, remaining)
    _low_stock_alert(p_name, len(remaining))
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame(to_send).to_excel(writer, index=False)
    out.seek(0)
    bot.send_document(uid, out, visible_file_name=f"{p_name}.xlsx",
                      caption=S.get("shop_success","✅ *Purchase Successful!*\nCost: *{total} BDT*").format(total=total),
                      parse_mode="Markdown")


def _deliver_vpn_manual(uid, p_name, p_info, qty, total, S, order_id=None):
    """Manual delivery — notify user to wait, alert admin to send VPN access."""
    d_name   = disp_name(p_info, p_name)
    duration = p_info.get("duration", "N/A")
    u        = get_user(uid)
    user_name = u.get("name") or u.get("first_name") or u.get("username") or "Unknown"
    username  = u.get("username") or "—"
    now      = bst_now()
    # Save order context — persisted to disk so it survives bot restarts
    _pending_manual_orders[uid] = {
        "d_name":    d_name,
        "duration":  duration,
        "total":     total,
        "qty":       qty,
        "time":      now,
        "user_name": user_name,
        "username":  username,
        "order_id":  order_id,
    }
    _save_manual_orders(_pending_manual_orders)
    # Notify user
    bot.send_message(
        uid,
        f"✅ *আপনার Purchase সফলভাবে সম্পন্ন হয়েছে!*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"🛡️ *Product:* {d_name}\n"
        f"⏳ *Duration:* {duration}\n"
        f"💸 *Total Paid:* *{total} BDT*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"🕐 অনুগ্রহ করে *১ থেকে ১০ মিনিট* অপেক্ষা করুন।\n"
        f"আপনার VPN Access শীঘ্রই পাঠিয়ে দেওয়া হবে। 🚀",
        parse_mode="Markdown"
    )
    # Alert admin for manual delivery
    try:
        mk_adm = types.InlineKeyboardMarkup()
        mk_adm.add(
            types.InlineKeyboardButton(
                "📨 Send VPN to User", callback_data=f"adm_send_vpn|{uid}"
            )
        )
        bot.send_message(
            ADMIN_ID,
            f"🔔 *Manual VPN Delivery Required!*\n"
            f"✨━━━━━━━━━━━━━━━━━━✨\n"
            f"👤 User ID: `{uid}`\n"
            f"🏷️ Name: {user_name}\n"
            f"📲 Username: @{username}\n"
            f"🛡️ Product: *{d_name}*\n"
            f"⏳ Duration: *{duration}*\n"
            f"🔢 Qty: *{qty}* | 💰 Total: *{total} BDT*\n"
            f"🕐 Time: {now}\n"
            f"✨━━━━━━━━━━━━━━━━━━✨\n"
            f"➡️ নিচের বাটনে ক্লিক করে VPN Access পাঠান।",
            reply_markup=mk_adm,
            parse_mode="Markdown"
        )
    except Exception:
        pass


@bot.callback_query_handler(func=lambda c: c.data.startswith("adm_send_vpn|"))
def adm_send_vpn_start(call):
    """Admin clicks 'Send VPN to User' — show order info, ask for email|password."""
    if call.message.chat.id != ADMIN_ID: return
    target_uid = call.data.split("|")[1]
    bot.answer_callback_query(call.id)
    order = _pending_manual_orders.get(target_uid, {})
    d_name   = order.get("d_name", "—")
    duration = order.get("duration", "—")
    total    = order.get("total", "—")
    u_bal    = round(get_user(target_uid).get("balance", 0), 2)
    msg = bot.send_message(
        ADMIN_ID,
        f"📨 *VPN Credentials পাঠান*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"👤 User: `{target_uid}`\n"
        f"🛡️ Product: *{d_name}*\n"
        f"⏳ Duration: *{duration}*\n"
        f"💸 Paid: *{total} BDT*\n"
        f"💰 User Balance Now: *{u_bal} BDT*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"এখন শুধু *email* এবং *password* এই ফরম্যাটে পাঠান:\n"
        f"`email@example.com|YourPassword123`\n\n"
        f"❌ বাতিল করতে লিখুন: `/cancel`",
        parse_mode="Markdown"
    )
    bot.register_next_step_handler(msg, _adm_send_vpn_deliver, target_uid)


def _adm_send_vpn_deliver(message, target_uid):
    """Admin sends email|password — auto-format and deliver to user."""
    if message.chat.id != ADMIN_ID: return
    if message.text and message.text.strip().lower() == "/cancel":
        bot.send_message(ADMIN_ID, "❌ বাতিল করা হয়েছে।")
        return
    try:
        cred_block = message.text.strip()

        order    = _pending_manual_orders.get(target_uid, {})
        d_name   = order.get("d_name", "—")
        duration = order.get("duration", "—")
        total    = order.get("total", "—")
        now      = bst_now()
        u_bal    = f"{round(get_user(target_uid).get('balance', 0), 2):.2f}"

        # Split "email|password" so each value is wrapped in its own `code`
        # block — tapping a `code` value in Telegram copies just that value.
        if "|" in cred_block:
            email_part, pass_part = cred_block.split("|", 1)
            cred_lines = (
                f"📧 *Email:*\n`{email_part.strip()}`\n\n"
                f"🔑 *Password:*\n`{pass_part.strip()}`"
            )
        else:
            cred_lines = f"`{cred_block}`"

        # Send the formatted VPN access to the user
        bot.send_message(
            target_uid,
            f"🛡️ *আপনার VPN Access এসে গেছে!*\n"
            f"✨━━━━━━━━━━━━━━━━━━✨\n"
            f"📦 *Product:* {d_name}\n"
            f"⏳ *Duration:* {duration}\n"
            f"💸 *Total Paid:* {total} BDT\n"
            f"💰 *Balance:* {u_bal} BDT\n"
            f"🕐 *Time:* {now}\n"
            f"✨━━━━━━━━━━━━━━━━━━✨\n"
            f"{cred_lines}\n"
            f"✨━━━━━━━━━━━━━━━━━━✨\n"
            f"👆 _উপরের Email/Password-এ ট্যাপ করলে কপি হয়ে যাবে_\n"
            f"🎉 *ধন্যবাদ! Prime Bazar ব্যবহার করার জন্য।*",
            parse_mode="Markdown"
        )
        # Confirm to admin
        bot.send_message(
            ADMIN_ID,
            f"✅ *সফলভাবে পাঠানো হয়েছে!*\n"
            f"👤 User `{target_uid}` VPN access পেয়েছেন।\n"
            f"🛡️ {d_name} | ⏳ {duration}",
            parse_mode="Markdown"
        )
        # Clean up pending order from memory and disk
        _pending_manual_orders.pop(target_uid, None)
        _save_manual_orders(_pending_manual_orders)
    except Exception as e:
        bot.send_message(
            ADMIN_ID,
            f"❌ *পাঠাতে সমস্যা হয়েছে!*\nError: `{e}`",
            parse_mode="Markdown"
        )


def _deliver_vpn(uid, call, p_name, qty, total, success_msg, S):
    """
    Deliver VPN credentials via userbot navigation only.
    No local xlsx fallback — VPN stock is managed by the supplier bot.
    Raises:
      ProductUnavailableError  – supplier has no matching product button
      SupplierTimeoutError     – supplier bot didn't respond within time limit
    """
    supplier = cfg("supplier_bot_username")
    if _userbot is None or not supplier:
        raise SupplierTimeoutError("Userbot not connected.")

    prods   = get_products()
    dur_key = prods[p_name].get("vpn_dur", "30d")

    note = bot.send_message(
        uid,
        "⏳ *Connecting to supplier...*\n"
        "🤖 Navigating menus, please wait a moment.",
        parse_mode="Markdown"
    )
    # Total outer timeout: 5 steps × 14s each × qty + buffer
    outer_timeout = max(120, qty * 80)
    result = _run_async(_vpn_buy_nav_async(p_name, dur_key, qty), timeout=outer_timeout)
    try: bot.delete_message(uid, note.message_id)
    except Exception: pass

    if result == "unavailable":
        raise ProductUnavailableError(f"Supplier bot does not have '{p_name}'.")

    if result == "timeout" or result is None:
        raise SupplierTimeoutError("Supplier bot did not respond in time.")

    # result is a list of (mail, password) tuples
    # Validate credentials — if all are empty dashes, supplier had no balance/stock
    valid = [(m, p) for m, p in result if m != "—" and p != "—"]
    if not valid:
        raise SupplierTimeoutError("Supplier returned empty credentials (likely insufficient balance).")

    delivery = success_msg + "\n\n"
    for i, (mail, pw) in enumerate(valid, 1):
        delivery += (
            f"🛡️ *VPN Account #{i}*\n"
            f"✨━━━━━━━━━━━━✨\n"
            f"📧 *Email:*\n`{mail}`\n\n"
            f"🔑 *Password:*\n`{pw}`\n"
            f"✨━━━━━━━━━━━━✨\n\n"
        )
    delivery += (
        f"👆 _Email বা Password-এ ট্যাপ করলে আলাদাভাবে কপি হবে_\n"
        f"🎉 *ধন্যবাদ! Prime Bazar ব্যবহার করার জন্য।* 🛒"
    )
    bot.send_message(uid, delivery.strip(), parse_mode="Markdown")
    # Check supplier balance after every VPN purchase (non-blocking)
    threading.Thread(target=_check_and_alert_balance_once, daemon=True).start()


def _check_and_alert_balance_once():
    """One-shot balance check — used after each VPN sale (no rescheduling)."""
    try:
        bal = _run_async(_check_supplier_balance_async(), timeout=25)
        if bal is not None:
            print(f"[BalCheck] Supplier balance after sale: {bal} BDT")
            if bal < _LOW_BAL_THRESHOLD:
                bot.send_message(ADMIN_ID,
                    f"⚠️ *সাপ্লায়ার বট ব্যালেন্স লো!*\n"
                    f"✨━━━━━━━━━━━━━━━━━━✨\n"
                    f"💰 বর্তমান ব্যালেন্স: *{bal} BDT*\n"
                    f"🔴 ব্যালেন্স ৫০ টাকার নিচে নেমেছে!\n"
                    f"✨━━━━━━━━━━━━━━━━━━✨\n"
                    f"📲 অনুগ্রহ করে দ্রুত রিচার্জ করুন।",
                    parse_mode="Markdown")
    except Exception as e:
        print(f"[BalCheck] one-shot error: {e}")


# ═══════════════════════════════════════════════
#  DEPOSIT
# ═══════════════════════════════════════════════

# Logo map — keys must match payment method names in settings.json
_PAYMENT_LOGO_FILES = {
    "Bkash":   "attached_assets/images_(1)_1782970910503.png",
    "Rocket":  "attached_assets/images_(2)_1782970910592.png",
    "Binance": "attached_assets/images_(11)_1782970910651.jpeg",
}

def _make_payment_banner(methods: list) -> io.BytesIO | None:
    """Build a horizontal banner showing circular logos + method names."""
    try:
        from PIL import ImageDraw, ImageFont
        LOGO  = 72          # diameter of each circle
        PAD_X = 30          # horizontal padding around each logo
        PAD_Y = 18          # top/bottom padding
        LBL_H = 22          # height reserved below each logo for label
        SLOT_W = LOGO + PAD_X * 2
        W = SLOT_W * len(methods)
        H = PAD_Y + LOGO + 8 + LBL_H + PAD_Y
        BG = (18, 18, 28)   # dark navy background

        banner = Image.new("RGBA", (W, H), BG + (255,))
        draw   = ImageDraw.Draw(banner)

        # Try to load a small system font; fall back to default
        try:
            font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 14)
        except Exception:
            font = ImageFont.load_default()

        for i, method in enumerate(methods):
            cx = i * SLOT_W + PAD_X   # left edge of this slot
            cy = PAD_Y                  # top edge

            # --- circular logo ---
            logo_path = _PAYMENT_LOGO_FILES.get(method)
            if logo_path:
                try:
                    logo = Image.open(logo_path).convert("RGBA")
                    logo = logo.resize((LOGO, LOGO), Image.LANCZOS)
                    # circular mask
                    mask = Image.new("L", (LOGO, LOGO), 0)
                    ImageDraw.Draw(mask).ellipse((0, 0, LOGO - 1, LOGO - 1), fill=255)
                    circ = Image.new("RGBA", (LOGO, LOGO), (0, 0, 0, 0))
                    circ.paste(logo, mask=mask)
                    banner.paste(circ, (cx, cy), circ)
                except Exception:
                    # fallback: coloured circle
                    ImageDraw.Draw(banner).ellipse(
                        (cx, cy, cx + LOGO - 1, cy + LOGO - 1), fill=(80, 80, 100))
            else:
                ImageDraw.Draw(banner).ellipse(
                    (cx, cy, cx + LOGO - 1, cy + LOGO - 1), fill=(80, 80, 100))

            # --- label ---
            lbl_y = cy + LOGO + 8
            lbl_x = cx + LOGO // 2
            draw.text((lbl_x, lbl_y), method, font=font, fill=(210, 210, 210), anchor="mt")

        out = io.BytesIO()
        banner.convert("RGB").save(out, format="PNG")
        out.seek(0)
        return out
    except Exception as e:
        print(f"Payment banner error: {e}")
        return None


def _adm_deposit_toggle(call):
    """Toggle deposit on/off and broadcast to all users."""
    current = cfg("deposit_enabled") is not False
    new_val  = not current
    settings["deposit_enabled"] = new_val
    save_settings(settings)
    state_txt = "চালু (ON) ✅" if new_val else "বন্ধ (OFF) ❌"
    bot.answer_callback_query(call.id, f"Deposit → {'ON ✅' if new_val else 'OFF ❌'}", show_alert=True)
    # Broadcast to all users
    if new_val:
        bcast = (
            "🟢 *ডিপোজিট সিস্টেম চালু হয়েছে!*\n"
            "✨━━━━━━━━━━━━━━━━━━✨\n"
            "💰 আপনি এখন ডিপোজিট করতে পারবেন।\n"
            "💳 Bkash, Rocket ও Binance এর মাধ্যমে সহজেই\n"
            "ব্যালেন্স যোগ করুন এবং কেনাকাটা উপভোগ করুন!\n"
            "✨━━━━━━━━━━━━━━━━━━✨\n"
            "🛒 *এখনই ডিপোজিট করুন এবং শপিং শুরু করুন!*"
        )
    else:
        bcast = (
            "🔴 *সাময়িকভাবে ডিপোজিট সিস্টেম বন্ধ রাখা হয়েছে।*\n"
            "✨━━━━━━━━━━━━━━━━━━✨\n"
            "⏳ অনুগ্রহ করে একটু অপেক্ষা করুন।\n"
            "🔔 সিস্টেম পুনরায় চালু হলে আপনাকে\n"
            "স্বয়ংক্রিয়ভাবে জানানো হবে।\n"
            "✨━━━━━━━━━━━━━━━━━━✨\n"
            "🙏 অসুবিধার জন্য আন্তরিকভাবে দুঃখিত।"
        )
    sent = failed = 0
    for uid in list(user_data.keys()):
        try:
            bot.send_message(uid, bcast, parse_mode="Markdown")
            sent += 1
        except Exception:
            failed += 1
    try:
        bot.edit_message_reply_markup(ADMIN_ID, call.message.message_id,
                                      reply_markup=admin_panel_markup())
    except Exception: pass
    bot.send_message(ADMIN_ID,
        f"✅ Deposit → *{'ON' if new_val else 'OFF'}*\n📨 Notified: {sent} | ❌ Failed: {failed}",
        parse_mode="Markdown")


@bot.message_handler(func=lambda m: m.text in _btn_labels("btn_deposit"))
def deposit_menu(message):
    if not guard(message): return
    if cfg("deposit_enabled") is False:
        uid = str(message.chat.id)
        S   = STRINGS.get(get_lang(uid), STRINGS["bn"])
        bot.send_message(message.chat.id,
            "🔴 *ডিপোজিট সিস্টেম বর্তমানে বন্ধ আছে।*\n"
            "✨━━━━━━━━━━━━━━━━━━✨\n"
            "⏳ শীঘ্রই চালু হবে। একটু অপেক্ষা করুন।",
            parse_mode="Markdown"); return
    uid     = str(message.chat.id)
    S       = STRINGS.get(get_lang(uid), STRINGS["bn"])
    methods = list(PAYMENT_METHODS().keys())
    mk      = types.InlineKeyboardMarkup(row_width=2)
    _method_icon = {"bkash": "📱", "rocket": "🚀", "nagad": "🟠", "binance": "🌐"}
    btns = [
        types.InlineKeyboardButton(
            f"{_method_icon.get(m.lower(), '💳')} {m}", callback_data=f"dep|{m}")
        for m in methods
    ]
    if btns: mk.add(*btns)
    caption = S.get("deposit_select", "💳 *Select a Payment Method*")
    bot.send_message(message.chat.id, caption, reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("dep|"))
def get_dep_amount(call):
    uid    = str(call.message.chat.id)
    S      = STRINGS.get(get_lang(uid), STRINGS["bn"])
    method = call.data.split("|")[1]
    min_bdt = cfg("min_deposit_bdt") or 20
    min_usd = cfg("min_deposit_usd") or 1
    if method == "Binance":
        msg = bot.send_message(call.message.chat.id,
            S.get("dep_ask_usd", "💵 *How much would you like to deposit?*\n🔹 Minimum: *${min_usd}*").format(min_usd=min_usd),
            parse_mode="Markdown")
    else:
        msg = bot.send_message(call.message.chat.id,
            S.get("dep_ask_bdt", "৳ *How much would you like to send?*\n🔹 Minimum: *{min_bdt} BDT*").format(min_bdt=min_bdt),
            parse_mode="Markdown")
    bot.register_next_step_handler(msg, process_deposit, method)
    bot.answer_callback_query(call.id)


def process_deposit(message, method):
    uid  = str(message.chat.id)
    S    = STRINGS.get(get_lang(uid), STRINGS["bn"])
    try: amount_input = float(message.text.strip())
    except ValueError:
        bot.send_message(message.chat.id, S.get("dep_invalid","❌ Invalid amount!"), parse_mode="Markdown"); return
    num = PAYMENT_METHODS().get(method, "")
    min_bdt = cfg("min_deposit_bdt") or 20
    min_usd = cfg("min_deposit_usd") or 1
    if method == "Binance":
        if amount_input < min_usd:
            bot.send_message(message.chat.id,
                S.get("dep_min_usd", "❌ *Minimum deposit via Binance is ${min_usd}.*").format(min_usd=min_usd),
                parse_mode="Markdown"); return
        bdt_amount = round(amount_input * USD_RATE(), 2)
        msg = bot.send_message(message.chat.id,
            S.get("dep_binance_info","✅ *Our Binance:* `{num}`\n💵 ${usd} = {bdt} BDT\n📩 Send Order ID:").format(
                num=num, usd=amount_input, bdt=bdt_amount), parse_mode="Markdown")
        bot.register_next_step_handler(msg, alert_admin, bdt_amount, method, amount_input)
    else:
        if amount_input < min_bdt:
            bot.send_message(message.chat.id,
                S.get("dep_min_bdt", "❌ *Minimum deposit is {min_bdt} BDT.*").format(min_bdt=min_bdt),
                parse_mode="Markdown"); return
        instructions = S.get(
            "dep_bdt_instructions",
            "🏦 *{method} Deposit*\n📲 Send to: `{num}`\n💰 Amount: *{amount} BDT*\n"
            "📩 Reply with your Transaction ID (TrxID) now:"
        ).format(method=method, num=num, amount=amount_input)
        mk = types.InlineKeyboardMarkup()
        mk.add(types.InlineKeyboardButton(
            S.get("dep_sent_btn", "✅ I've Sent the Money"), callback_data="depsent"))
        msg = bot.send_message(message.chat.id, instructions, reply_markup=mk, parse_mode="Markdown")
        bot.register_next_step_handler(msg, alert_admin, amount_input, method)


@bot.callback_query_handler(func=lambda c: c.data == "dep_start")
def dep_start_retry(call):
    """'Try Again' button after an invalid TrxID — re-show payment methods."""
    uid     = str(call.message.chat.id)
    S       = STRINGS.get(get_lang(uid), STRINGS["bn"])
    methods = list(PAYMENT_METHODS().keys())
    _method_icon = {"bkash": "📱", "rocket": "🚀", "nagad": "🟠", "binance": "🌐"}
    mk      = types.InlineKeyboardMarkup(row_width=2)
    btns    = [types.InlineKeyboardButton(f"{_method_icon.get(m.lower(), '💳')} {m}", callback_data=f"dep|{m}")
               for m in methods]
    if btns: mk.add(*btns)
    bot.answer_callback_query(call.id)
    bot.send_message(call.message.chat.id,
        S.get("deposit_select", "💳 *Select a Payment Method*"), reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data == "depsent")
def dep_sent_ack(call):
    """Cosmetic 'I've sent the money' suggestion button — reinforces the next step."""
    uid = str(call.message.chat.id)
    S   = STRINGS.get(get_lang(uid), STRINGS["bn"])
    bot.answer_callback_query(call.id)
    bot.send_message(call.message.chat.id,
        S.get("dep_sent_ack", "📝 *Great!* Please type your Transaction ID (TrxID) now."),
        parse_mode="Markdown")


def _validate_trxid(trx: str, method: str):
    """Return (is_valid, error_message). Validates TrxID format per payment method."""
    import re
    trx = trx.strip()
    if not trx:
        return False, "❌ *TrxID খালি।* সঠিক Transaction ID পাঠান।"
    if " " in trx:
        return False, "❌ *TrxID-এ স্পেস থাকতে পারবে না।*\nTransaction ID-তে কোনো স্পেস বা ফাঁক দেবেন না।"
    if method in ("Bkash", "Rocket"):
        if not re.match(r'^[A-Za-z0-9]+$', trx):
            return False, (
                "❌ *TrxID সঠিক নয়।*\n"
                f"{method} TrxID শুধুমাত্র *ইংরেজি অক্ষর* (A-Z) ও *সংখ্যা* (0-9) দিয়ে হয়।\n"
                "বিশেষ চিহ্ন, বাংলা বা স্পেস গ্রহণযোগ্য নয়।"
            )
        if len(trx) < 8 or len(trx) > 20:
            return False, (
                "❌ *TrxID Trx Not Found।*\n"
                f"{method} Transaction ID সাধারণত ৮–২০ অক্ষরের হয়।\n"
                "আপনার TrxID পুনরায় চেক করুন এবং সঠিকটি পাঠান।"
            )
    elif method == "Binance":
        if not re.match(r'^[A-Za-z0-9\-]+$', trx):
            return False, (
                "❌ *Order ID সঠিক নয়।*\n"
                "Binance Order ID শুধুমাত্র সংখ্যা ও অক্ষর দিয়ে হয়।"
            )
        if len(trx) < 6:
            return False, "❌ *Binance Order ID সঠিক নয়।* পুনরায় চেক করুন।"
    else:
        if len(trx) < 5:
            return False, "❌ *TrxID অনেক ছোট।* সঠিক Transaction ID পাঠান।"
    return True, ""


def alert_admin(message, bdt_amount, method, usd_amount=None):
    global used_trxids
    uid  = str(message.chat.id)
    S    = STRINGS.get(get_lang(uid), STRINGS["bn"])
    trx  = message.text.strip() if message.text else ""

    # ── Format Validation ───────────────────────
    valid, err_msg = _validate_trxid(trx, method)
    if not valid:
        mk_retry = types.InlineKeyboardMarkup()
        mk_retry.add(types.InlineKeyboardButton(
            "🔄 Try Again" if get_lang(uid) == "en" else "🔄 আবার চেষ্টা করুন", callback_data="dep_start"))
        bot.send_message(message.chat.id, err_msg, parse_mode="Markdown", reply_markup=mk_retry)
        return

    # ── Duplicate TrxID check ────────────────────
    trx_key = trx.lower().strip()
    if trx_key in used_trxids:
        dup_msg = (
            "🚫 *This TrxID has already been used!*\n"
            "✨━━━━━━━━━━━━━━━━━━✨\n"
            "❌ Duplicate Transaction IDs are not accepted.\n"
            "If this is a mistake, please contact Support.\n"
            "✨━━━━━━━━━━━━━━━━━━✨"
        ) if get_lang(uid) == "en" else (
            "🚫 *এই TrxID আগে ব্যবহার করা হয়েছে!*\n"
            "✨━━━━━━━━━━━━━━━━━━✨\n"
            "❌ Duplicate Transaction ID গ্রহণযোগ্য নয়।\n"
            "সঠিক TrxID না থাকলে Support-এ যোগাযোগ করুন।\n"
            "✨━━━━━━━━━━━━━━━━━━✨"
        )
        bot.send_message(message.chat.id, dup_msg, parse_mode="Markdown"); return

    mk = types.InlineKeyboardMarkup(row_width=2)
    mk.add(
        types.InlineKeyboardButton("✅ Approve", callback_data=f"adm_aprv_{uid}_{int(bdt_amount)}_{trx_key[:20]}"),
        types.InlineKeyboardButton("❌ Reject",  callback_data=f"adm_rej_{uid}"),
    )
    if method == "Binance":
        admin_txt = (
            f"🔔 *নতুন ডিপোজিট (Binance)*\n✨━━━━━━━━━━━━✨\n"
            f"👤 ইউজার: `{uid}`\n💵 ${usd_amount} = {bdt_amount} BDT\n"
            f"💳 মেথড: {method}\n📋 Info: `{trx}`\n✨━━━━━━━━━━━━✨"
        )
    else:
        admin_txt = (
            f"🔔 *নতুন ডিপোজিট*\n✨━━━━━━━━━━━━✨\n"
            f"👤 ইউজার: `{uid}`\n💰 {bdt_amount} BDT\n"
            f"💳 মেথড: {method}\n📋 TrxID: `{trx}`\n✨━━━━━━━━━━━━✨"
        )
    bot.send_message(ADMIN_ID, admin_txt, parse_mode="Markdown", reply_markup=mk)
    # Save to pending deposits so admin can list them
    pending_deposits[uid] = {
        "uid": uid, "method": method, "bdt": bdt_amount,
        "usd": usd_amount, "trx": trx, "time": bst_now(),
        "trx_key": trx_key,
    }
    _save_pending_deps(pending_deposits)
    support_url = f"https://t.me/{SUPPORT_USERNAME().lstrip('@')}" if SUPPORT_USERNAME() else "https://t.me/owner_of_pam"
    mk_sup = types.InlineKeyboardMarkup()
    mk_sup.add(types.InlineKeyboardButton("🆘 Support", url=support_url))
    bot.send_message(message.chat.id,
        S.get("dep_sent",
            "✅ *Your deposit request has been submitted!*\n"
            "✨━━━━━━━━━━━━━━━━━━✨\n"
            "⏳ Please wait a few minutes.\n\n"
            "🕐 *If your balance isn't credited within 1–10 minutes*\n"
            "please contact Support.\n"
            "✨━━━━━━━━━━━━━━━━━━✨"),
        reply_markup=mk_sup,
        parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("adm_aprv_") or c.data.startswith("adm_rej_"))
def admin_deposit_decision(call):
    global used_trxids, pending_deposits, rejected_deposits
    if call.message.chat.id != ADMIN_ID: return
    parts = call.data.split("_")
    if call.data.startswith("adm_aprv_"):
        # format: adm_aprv_{uid}_{amt}_{trx_key}  (trx_key may be absent for old requests)
        tid = parts[2]; amt = int(parts[3])
        trx_key = parts[4] if len(parts) > 4 else None
        update_user(tid, "balance", amt); update_user(tid, "total_deposit", amt)
        user_data[tid]["last_deposit_time"] = bst_now()
        dep_rec = {"amount": amt, "date": str(bst_now())}
        if "deposit_history" not in user_data[tid]:
            user_data[tid]["deposit_history"] = []
        user_data[tid]["deposit_history"].append(dep_rec)
        save_data(user_data)
        # Save TrxID so it can't be reused
        if trx_key:
            used_trxids.add(trx_key)
            save_trxids(used_trxids)
        # Remove from pending
        pending_deposits.pop(tid, None); _save_pending_deps(pending_deposits)
        S = STRINGS.get(get_lang(tid), STRINGS["bn"])
        bot.send_message(tid, S.get("dep_approved","🎉 *{amt} BDT* added!").format(amt=amt), parse_mode="Markdown")
        bot.edit_message_text(f"✅ Approved {amt} BDT → `{tid}`",
                              call.message.chat.id, call.message.message_id, parse_mode="Markdown")
    else:
        tid  = parts[2]
        dep  = pending_deposits.pop(tid, {})
        if dep:
            rejected_deposits.append({**dep, "rejected_at": bst_now()})
            if len(rejected_deposits) > 200: rejected_deposits = rejected_deposits[-200:]
            _save_rejected_deps(rejected_deposits)
            _save_pending_deps(pending_deposits)
        S    = STRINGS.get(get_lang(tid), STRINGS["bn"])
        bot.send_message(tid, S.get("dep_rejected","❌ Deposit rejected."), parse_mode="Markdown")
        bot.edit_message_text(f"❌ Rejected deposit for `{tid}`",
                              call.message.chat.id, call.message.message_id, parse_mode="Markdown")


# ═══════════════════════════════════════════════
#  DAILY BONUS
# ═══════════════════════════════════════════════
@bot.message_handler(func=lambda m: m.text in _btn_labels("btn_daily"))
def daily_bonus(message):
    if not guard(message): return
    if not cfg("daily_bonus_enabled"): return
    uid  = str(message.chat.id)
    S    = STRINGS.get(get_lang(uid), STRINGS["bn"])
    u    = get_user(uid)
    last = u.get("last_daily_bonus"); now = datetime.utcnow()
    if last:
        try:
            diff = (now - datetime.fromisoformat(last)).total_seconds()
            if diff < 86400:
                wait_h = int((86400-diff)//3600); wait_m = int(((86400-diff)%3600)//60)
                bot.send_message(message.chat.id,
                    S.get("daily_wait","⏰ Next bonus in {hrs}h {mins}m.").format(hrs=wait_h, mins=wait_m),
                    parse_mode="Markdown"); return
        except Exception: pass
    amt = cfg("daily_bonus_amount")
    user_data[uid]["balance"] = round(u.get("balance", 0) + amt, 2)
    user_data[uid]["last_daily_bonus"] = now.isoformat(); save_data(user_data)
    bot.send_message(message.chat.id,
        S.get("daily_claimed","🎁 *Daily Bonus!*\n💰 *{amt} BDT* added!").format(amt=amt),
        parse_mode="Markdown")


# ═══════════════════════════════════════════════
#  ORDERS / COUPON
# ═══════════════════════════════════════════════
def _prompt_review(uid, order_id, p_name, S):
    """Ask the buyer to rate their just-completed order with a 1–5 star tap.
    Only shown on the very first purchase — avoids annoying repeat buyers."""
    try:
        if get_user(uid).get("total_orders", 0) != 1:
            return
        mk = types.InlineKeyboardMarkup(row_width=5)
        mk.add(*[
            types.InlineKeyboardButton(f"{'⭐' * n}", callback_data=f"rate|{order_id}|{n}")
            for n in range(1, 6)
        ])
        bot.send_message(
            uid,
            S.get("review_prompt", "🌟 *How was your experience with {product}?*\nTap to rate:").format(product=p_name),
            reply_markup=mk, parse_mode="Markdown")
    except Exception:
        pass


@bot.callback_query_handler(func=lambda c: c.data.startswith("rate|"))
def cb_submit_review(call):
    """User tapped a star rating — store it and optionally ask for a comment."""
    uid  = str(call.message.chat.id)
    S    = STRINGS.get(get_lang(uid), STRINGS["bn"])
    _, order_id, stars = call.data.split("|")
    stars = int(stars)
    bot.answer_callback_query(call.id, "✅ Thanks for your feedback!")
    # Find product name from the user's own order history for context
    order = next((o for o in get_user(uid).get("orders", []) if o.get("order_id") == order_id), {})
    product = order.get("product", "—")
    db_append_review({
        "order_id": order_id, "uid": uid, "product": product,
        "rating": stars, "date": bst_now(),
    })
    try:
        bot.edit_message_text(
            f"{'⭐' * stars}\n" + S.get("review_thanks", "✅ *Thank you for rating this order!*"),
            call.message.chat.id, call.message.message_id, parse_mode="Markdown")
    except Exception: pass


@bot.message_handler(commands=["reviews"])
def admin_reviews(message):
    """Quick command: shows average rating and the most recent reviews."""
    if not staff_only(message): return
    reviews = db_load_reviews()
    if not reviews:
        bot.send_message(message.chat.id, "📭 এখনো কোনো রিভিউ আসেনি।"); return
    avg = round(sum(r.get("rating", 0) for r in reviews) / len(reviews), 2)
    stars_summary = {n: sum(1 for r in reviews if r.get("rating") == n) for n in range(1, 6)}
    lines = [
        "⭐ *Customer Reviews*",
        "✨━━━━━━━━━━━━━━━━━━✨",
        f"📊 Average Rating: *{avg} / 5*  ({len(reviews)} reviews)",
    ]
    for n in range(5, 0, -1):
        lines.append(f"{'⭐' * n}: {stars_summary[n]}")
    lines.append("✨━━━━━━━━━━━━━━━━━━✨")
    lines.append("*Recent reviews:*")
    for r in reversed(reviews[-8:]):
        lines.append(f"{'⭐' * r.get('rating', 0)} · {r.get('product','—')} · {r.get('date','—')}")
    bot.send_message(message.chat.id, "\n".join(lines), parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data == "order_history")
def cb_order_history(call):
    uid  = str(call.message.chat.id)
    S    = STRINGS.get(get_lang(uid), STRINGS["bn"])
    u    = get_user(uid)
    orders = u.get("orders", [])
    bot.answer_callback_query(call.id)
    if not orders:
        bot.send_message(uid, S.get("orders_empty","📭 No orders yet."), parse_mode="Markdown"); return
    txt = S.get("orders_title","🛒 *ORDER HISTORY*") + "\n✨━━━━━━━━━━━━━━━━━━✨\n"
    for i, o in enumerate(reversed(orders[-10:]), 1):
        txt += f"*{i}.* {o['product']} × {o['qty']} — *{o['total']} BDT*\n    📅 {o['date']}\n\n"
    txt += S.get("orders_footer","✨━━━━━━━━━━━━✨\n📊 Total: *{count}*").format(count=len(orders))
    bot.send_message(uid, txt, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data == "deposit_history")
def cb_deposit_history(call):
    uid  = str(call.message.chat.id)
    lang = get_lang(uid)
    u    = get_user(uid)
    hist = u.get("deposit_history", [])
    bot.answer_callback_query(call.id)
    if not hist:
        msg = "💳 *Deposit History*\n✨━━━━━━━━━━━━━━━━━━✨\n📭 কোনো ডিপোজিট রেকর্ড নেই।" \
              if lang == "bn" else "💳 *Deposit History*\n✨━━━━━━━━━━━━━━━━━━✨\n📭 No deposit records yet."
        bot.send_message(uid, msg, parse_mode="Markdown"); return
    txt = "💳 *ডিপোজিট হিস্টোরি*\n✨━━━━━━━━━━━━━━━━━━✨\n" \
          if lang == "bn" else "💳 *Deposit History*\n✨━━━━━━━━━━━━━━━━━━✨\n"
    for i, d in enumerate(reversed(hist[-15:]), 1):
        txt += f"*{i}.* 💰 *{d['amount']} BDT*\n    📅 {d['date']}\n\n"
    total = sum(d["amount"] for d in hist)
    txt += f"✨━━━━━━━━━━━━━━━━━━✨\n📊 মোট ডিপোজিট: *{total} BDT*"
    bot.send_message(uid, txt, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data == "resend_last_order")
def cb_resend_last_order(call):
    uid  = str(call.message.chat.id)
    lang = get_lang(uid)
    S    = STRINGS.get(lang, STRINGS["bn"])
    u    = get_user(uid)
    orders = u.get("orders", [])
    bot.answer_callback_query(call.id)
    if not orders:
        bot.send_message(uid, S.get("orders_empty","📭 No orders yet."), parse_mode="Markdown"); return
    last   = orders[-1]
    p_name = last.get("product","")
    qty    = last.get("qty", 1)
    total  = last.get("total", 0)
    date   = last.get("date","—")
    prods  = get_products()
    stock  = get_stock_count(p_name) if p_name in prods else 0
    price  = prods[p_name]["price"] if p_name in prods else 0
    new_total = round(qty * price, 2)
    cat    = prods[p_name].get("cat","") if p_name in prods else ""
    txt = (
        f"🔄 *Resend Last Order*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"📦 *Product:* {p_name}\n"
        f"🔢 *Qty:* {qty}\n"
        f"💸 *New Total:* {new_total} BDT\n"
        f"💰 *Your Balance:* {u['balance']:.2f} BDT\n"
        f"📅 *Last Ordered:* {date}\n"
        f"📦 *Stock:* {stock} pcs\n"
        f"✨━━━━━━━━━━━━━━━━━━✨"
    )
    mk = types.InlineKeyboardMarkup(row_width=1)
    if p_name in prods and stock >= qty:
        if cat == "mail":
            mode = "retail" if qty <= 6 else "bulk"
            mk.add(types.InlineKeyboardButton(
                S.get("shop_confirm","✅ Confirm Purchase"),
                callback_data=f"mailpay|{p_name}|{qty}|{mode}"))
        elif cat == "vpn":
            mk.add(types.InlineKeyboardButton(
                S.get("shop_confirm","✅ Confirm Purchase"),
                callback_data=f"pay|{p_name}|{qty}"))
    else:
        txt += "\n\n❌ *Product unavailable or out of stock.*"
    mk.add(types.InlineKeyboardButton("❌ Cancel", callback_data="noop"))
    bot.send_message(uid, txt, reply_markup=mk, parse_mode="Markdown")


@bot.message_handler(commands=["coupon"])
def coupon_start(message):
    if not guard(message): return
    S   = STRINGS.get(get_lang(message.chat.id), STRINGS["bn"])
    msg = bot.send_message(message.chat.id, S.get("coupon_prompt","🎟️ Enter your coupon code:"), parse_mode="Markdown")
    bot.register_next_step_handler(msg, coupon_apply)


def coupon_apply(message):
    uid  = str(message.chat.id)
    S    = STRINGS.get(get_lang(uid), STRINGS["bn"])
    code = message.text.strip().upper()
    global coupons; coupons = load_coupons()
    if code not in coupons:
        bot.send_message(message.chat.id, S.get("coupon_invalid","❌ Invalid coupon."), parse_mode="Markdown"); return
    c = coupons[code]
    if uid in [str(x) for x in c.get("used_by", [])]:
        bot.send_message(message.chat.id, S.get("coupon_used","❌ Already used."), parse_mode="Markdown"); return
    if c.get("uses_left", 0) <= 0:
        bot.send_message(message.chat.id, S.get("coupon_invalid","❌ Coupon expired."), parse_mode="Markdown"); return
    amt = c.get("amount", 0)
    user_data[uid]["balance"] = round(get_user(uid).get("balance", 0) + amt, 2); save_data(user_data)
    coupons[code]["uses_left"] -= 1; coupons[code].setdefault("used_by",[]).append(uid); save_coupons(coupons)
    bot.send_message(message.chat.id,
        S.get("coupon_ok","🎉 *{amt} BDT* added!").format(amt=amt), parse_mode="Markdown")


# ═══════════════════════════════════════════════════════════════
#  ADMIN PANEL
# ═══════════════════════════════════════════════════════════════
def admin_only(obj):
    uid = obj.from_user.id if hasattr(obj, "from_user") else obj.chat.id
    return uid == ADMIN_ID


def is_staff(chat_id) -> bool:
    """Owner (ADMIN_ID) is always staff. Staff members are extra Telegram IDs
    the owner has granted limited access to (order fulfillment + stock only)."""
    try:
        cid = int(chat_id)
    except (TypeError, ValueError):
        return False
    if cid == ADMIN_ID:
        return True
    return str(cid) in {str(s) for s in (cfg("staff_ids") or [])}


def staff_only(obj) -> bool:
    uid = obj.from_user.id if hasattr(obj, "from_user") else obj.chat.id
    return is_staff(uid)

_ADMIN_PANEL_TXT = (
    "🔑 *MASTER ADMIN PANEL*\n"
    "✨━━━━━━━━━━━━━━━━━━✨\n"
    "👋 Welcome, Admin!\n"
    "✨━━━━━━━━━━━━━━━━━━✨"
)

def _adm_section(mk, title):
    """Full-width, non-clickable section header for visual grouping."""
    mk.add(types.InlineKeyboardButton(f"▬▬▬▬▬  {title}  ▬▬▬▬▬", callback_data="adm|noop"))


def admin_panel_markup():
    mk = types.InlineKeyboardMarkup(row_width=2)

    _adm_section(mk, "🛍️ CATALOG")
    mk.add(
        types.InlineKeyboardButton("📦 Categories",      callback_data="adm|cats"),
        types.InlineKeyboardButton("🛍️ Products",        callback_data="adm|prods"),
    )
    mk.add(
        types.InlineKeyboardButton("🛡️ VPN Durations",   callback_data="adm|vpndurs"),
        types.InlineKeyboardButton("🎟️ Coupons",         callback_data="adm|coupons"),
    )

    _adm_section(mk, "📦 STOCK")
    mk.add(
        types.InlineKeyboardButton("➕ Add Stock",       callback_data="adm|addstock"),
        types.InlineKeyboardButton("🔄 Sync Stock",      callback_data="adm|syncstock"),
    )
    mk.add(
        types.InlineKeyboardButton("✏️ Stock Manager",   callback_data="adm|stockmgr"),
        types.InlineKeyboardButton("📊 Mail Report",     callback_data="adm|mailreport"),
    )

    _adm_section(mk, "👥 USERS & ORDERS")
    manual_cnt = len(_pending_manual_orders)
    mk.add(
        types.InlineKeyboardButton("👥 User Manager",    callback_data="adm|users"),
        types.InlineKeyboardButton(
            f"📋 Manual Orders ({manual_cnt})" if manual_cnt else "📋 Manual Orders",
            callback_data="adm|manual_orders"
        ),
    )

    _adm_section(mk, "📊 ANALYTICS & SETTINGS")
    mk.add(
        types.InlineKeyboardButton("📊 Stats",           callback_data="adm|stats"),
        types.InlineKeyboardButton("📈 Top Buyers",      callback_data="adm|topbuyers"),
    )
    mk.add(
        types.InlineKeyboardButton("🎨 Branding",        callback_data="adm|branding"),
        types.InlineKeyboardButton("❓ FAQ Manager",     callback_data="adm|faqmgr"),
    )
    mk.add(
        types.InlineKeyboardButton("🔘 Button Labels",   callback_data="adm|btnlabels"),
        types.InlineKeyboardButton("⚙️ Global Config",   callback_data="adm|config"),
    )
    mk.add(
        types.InlineKeyboardButton("🔀 Feature Toggles", callback_data="adm|toggles"),
        types.InlineKeyboardButton("🤖 Userbot Config",  callback_data="adm|userbot"),
    )

    _adm_section(mk, "👮 SUB ADMIN")
    mk.add(
        types.InlineKeyboardButton("👤 Sub Admins",      callback_data="adm|subadmins"),
    )

    _adm_section(mk, "⚙️ SYSTEM")
    dep_on = cfg("deposit_enabled") is not False
    mk.add(
        types.InlineKeyboardButton("📢 Broadcast",       callback_data="adm|broadcast"),
        types.InlineKeyboardButton(f"{'✅' if dep_on else '❌'} Deposit System",
                                   callback_data="adm|dep_toggle"),
    )
    mk.add(
        types.InlineKeyboardButton("💾 Backup DB",       callback_data="adm|backup"),
        types.InlineKeyboardButton("📥 Restore DB",      callback_data="adm|restore"),
    )
    if _REPLIT_DOMAIN:
        mk.add(
            types.InlineKeyboardButton("🌐 Web Admin Panel",
                                       url=f"https://{_REPLIT_DOMAIN}/admin/"),
        )
    return mk


@bot.message_handler(commands=["admin"])
def admin_panel(message):
    if not admin_only(message): return
    bot.send_message(ADMIN_ID, _ADMIN_PANEL_TXT,
                     reply_markup=admin_panel_markup(), parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("adm|"))
def admin_router(call):
    if call.message.chat.id != ADMIN_ID: return
    action = call.data.split("|")[1]
    if action == "back":
        bot.answer_callback_query(call.id)
        try:
            bot.edit_message_text(_ADMIN_PANEL_TXT, ADMIN_ID, call.message.message_id,
                                  reply_markup=admin_panel_markup(), parse_mode="Markdown")
        except Exception: pass
        return
    bot.answer_callback_query(call.id)
    {
        "stats":     _adm_stats,
        "cats":      _adm_cats,
        "prods":     _adm_prods_cat_select,
        "vpndurs":   _adm_vpndurs,
        "users":     _adm_usermgr,
        "branding":  _adm_branding,
        "config":    _adm_config,
        "toggles":   _adm_toggles,
        "coupons":       _adm_coupons,
        "userbot":       _adm_userbot,
        "btnlabels":     _adm_btnlabels,
        "manual_orders": _adm_manual_orders,
        "mailreport":    _adm_mailreport_select,
        "subadmins":     _adm_subadmins,
        "stockmgr":      _adm_stockmgr,
        "topbuyers":     _adm_top_buyers,
        "faqmgr":        _adm_faqmgr,
    }.get(action, lambda _: None)(call)

    if action == "addstock":
        _adm_addstock_select(call.message)
    elif action == "syncstock":
        _adm_syncstock(call)
    elif action == "dep_toggle":
        _adm_deposit_toggle(call)
    elif action == "broadcast":
        msg = bot.send_message(ADMIN_ID, "📢 সব ইউজারকে কোন মেসেজটি পাঠাতে চান? লিখুন:")
        bot.register_next_step_handler(msg, do_broadcast)
    elif action == "backup":
        _do_backup()
    elif action == "restore":
        msg = bot.send_message(ADMIN_ID, "📥 ইউজার ডাটা (.json) ফাইলটি পাঠান।")
        bot.register_next_step_handler(msg, do_restore)


# ─────────────────────────────────────────────
#  Sub Admin Management
# ─────────────────────────────────────────────
def _adm_subadmins(call):
    """Show current sub-admins with add/remove buttons."""
    staff = list(cfg("staff_ids") or [])
    mk = types.InlineKeyboardMarkup(row_width=1)
    if staff:
        for sid in staff:
            mk.add(types.InlineKeyboardButton(
                f"❌ Remove: {sid}", callback_data=f"subadm|remove|{sid}"))
    mk.add(types.InlineKeyboardButton("➕ নতুন Sub Admin যোগ করুন", callback_data="subadm|add"))
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="adm|back"))
    txt = (
        "👮 *Sub Admin Management*\n"
        "✨━━━━━━━━━━━━━━━━━━✨\n"
        + (("\n".join(f"• `{sid}`" for sid in staff)) if staff else "⚠️ কোনো Sub Admin নেই।")
        + "\n✨━━━━━━━━━━━━━━━━━━✨\n"
        "➕ যোগ করতে নিচের বাটন চাপুন।\n"
        "❌ Remove করতে ID-র পাশের বাটন চাপুন।"
    )
    try:
        bot.edit_message_text(txt, ADMIN_ID, call.message.message_id,
                              reply_markup=mk, parse_mode="Markdown")
    except Exception:
        bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("subadm|"))
def subadmin_router(call):
    if call.message.chat.id != ADMIN_ID: return
    parts  = call.data.split("|")
    action = parts[1]
    bot.answer_callback_query(call.id)

    if action == "add":
        msg = bot.send_message(
            ADMIN_ID,
            "👤 নতুন Sub Admin-এর *Telegram ID* দিন (সংখ্যা, যেমন: `123456789`):",
            parse_mode="Markdown")
        bot.register_next_step_handler(msg, _subadm_save_new)

    elif action == "remove" and len(parts) > 2:
        remove_id = parts[2]
        staff = list(cfg("staff_ids") or [])
        staff = [s for s in staff if str(s) != str(remove_id)]
        settings["staff_ids"] = staff
        save_settings(settings)
        bot.send_message(ADMIN_ID,
            f"✅ Sub Admin `{remove_id}` সরিয়ে দেওয়া হয়েছে।", parse_mode="Markdown")
        _adm_subadmins(call)


def _subadm_save_new(message):
    if message.chat.id != ADMIN_ID: return
    new_id = message.text.strip()
    if not new_id.lstrip("-").isdigit():
        bot.send_message(ADMIN_ID, "❌ Valid Telegram ID দিন (শুধু সংখ্যা)।"); return
    new_id = int(new_id)
    if new_id == ADMIN_ID:
        bot.send_message(ADMIN_ID, "❌ আপনি নিজেই Master Admin — এটা add করার দরকার নেই।"); return
    staff = list(cfg("staff_ids") or [])
    if str(new_id) in {str(s) for s in staff}:
        bot.send_message(ADMIN_ID, f"⚠️ `{new_id}` ইতিমধ্যে Sub Admin হিসেবে আছে।",
                         parse_mode="Markdown"); return
    staff.append(new_id)
    settings["staff_ids"] = staff
    save_settings(settings)
    bot.send_message(ADMIN_ID,
        f"✅ `{new_id}` কে Sub Admin হিসেবে যোগ করা হয়েছে!\n"
        f"এখন তিনি স্টাফ-অ্যাক্সেস পাবেন।",
        parse_mode="Markdown")


# ─────────────────────────────────────────────
#  Top Buyers
# ─────────────────────────────────────────────
def _adm_top_buyers(call):
    top = sorted(
        [(uid, u) for uid, u in user_data.items() if u.get("total_orders", 0) > 0],
        key=lambda x: x[1].get("total_orders", 0), reverse=True
    )[:10]
    mk = types.InlineKeyboardMarkup()
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="adm|back"))
    if not top:
        txt = "📈 *Top Buyers*\n✨━━━━━━━━━━━━━━━━━━✨\n⚠️ এখনো কোনো অর্ডার নেই।"
    else:
        txt = "📈 *Top Buyers — সেরা ১০ কাস্টমার*\n✨━━━━━━━━━━━━━━━━━━✨\n"
        medals = ["🥇", "🥈", "🥉"] + ["🏅"] * 7
        for i, (uid, u) in enumerate(top):
            fn   = u.get("first_name", "") or u.get("username", "") or uid
            ords = u.get("total_orders", 0)
            dep  = u.get("total_deposit", 0)
            txt += f"{medals[i]} *{fn}* (`{uid}`)\n   🛒 {ords} অর্ডার  💰 {dep} BDT\n\n"
        txt += "✨━━━━━━━━━━━━━━━━━━✨"
    try:
        bot.edit_message_text(txt, ADMIN_ID, call.message.message_id,
                              reply_markup=mk, parse_mode="Markdown")
    except Exception:
        bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")


# ─────────────────────────────────────────────
#  FAQ Manager (Admin)
# ─────────────────────────────────────────────
def _adm_faqmgr(call):
    items  = cfg("faq_items") or []
    enabled = cfg("faq_enabled") is True
    mk = types.InlineKeyboardMarkup(row_width=1)
    mk.add(types.InlineKeyboardButton(
        f"{'✅ FAQ চালু আছে' if enabled else '❌ FAQ বন্ধ আছে'} — টগল করুন",
        callback_data="faqadm|toggle"))
    for i, item in enumerate(items):
        mk.add(types.InlineKeyboardButton(
            f"❌ [{i+1}] {item['keyword'][:30]}",
            callback_data=f"faqadm|del|{i}"))
    mk.add(types.InlineKeyboardButton("➕ নতুন FAQ যোগ করুন", callback_data="faqadm|add"))
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="adm|back"))
    txt = (
        "❓ *FAQ Manager*\n"
        "✨━━━━━━━━━━━━━━━━━━✨\n"
        f"📋 মোট FAQ: *{len(items)}*\n"
        f"স্ট্যাটাস: {'✅ চালু' if enabled else '❌ বন্ধ'}\n"
        "✨━━━━━━━━━━━━━━━━━━✨\n"
        "কেউ keyword লিখলে অটো উত্তর যাবে।\n"
        "❌ বাটন চাপলে সেই FAQ মুছে যাবে।"
    )
    try:
        bot.edit_message_text(txt, ADMIN_ID, call.message.message_id,
                              reply_markup=mk, parse_mode="Markdown")
    except Exception:
        bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("faqadm|"))
def faqadm_router(call):
    if call.message.chat.id != ADMIN_ID: return
    parts  = call.data.split("|", 2)
    action = parts[1]
    bot.answer_callback_query(call.id)

    if action == "toggle":
        new_val = not (cfg("faq_enabled") is True)
        settings["faq_enabled"] = new_val
        save_settings(settings)
        _adm_faqmgr(call)

    elif action == "add":
        msg = bot.send_message(
            ADMIN_ID,
            "❓ *নতুন FAQ যোগ করুন*\n\n"
            "এই ফরম্যাটে দুই লাইনে লিখুন:\n"
            "`keyword`\n`উত্তর টেক্সট`\n\n"
            "উদাহরণ:\n`দাম\n১ মাস Netflix ২৫০ টাকা`",
            parse_mode="Markdown")
        bot.register_next_step_handler(msg, _faqadm_save)

    elif action == "del" and len(parts) > 2:
        idx = int(parts[2])
        items = list(cfg("faq_items") or [])
        if 0 <= idx < len(items):
            removed = items.pop(idx)
            settings["faq_items"] = items
            save_settings(settings)
            bot.send_message(ADMIN_ID,
                f"✅ FAQ মুছে দেওয়া হয়েছে: `{removed['keyword']}`",
                parse_mode="Markdown")
        _adm_faqmgr(call)


def _faqadm_save(message):
    if message.chat.id != ADMIN_ID: return
    lines = message.text.strip().split("\n", 1)
    if len(lines) < 2 or not lines[0].strip() or not lines[1].strip():
        bot.send_message(ADMIN_ID,
            "❌ দুই লাইনে লিখুন:\n`keyword`\n`উত্তর`", parse_mode="Markdown"); return
    keyword = lines[0].strip().lower()
    answer  = lines[1].strip()
    items   = list(cfg("faq_items") or [])
    items.append({"keyword": keyword, "answer": answer})
    settings["faq_items"] = items
    save_settings(settings)
    bot.send_message(ADMIN_ID,
        f"✅ FAQ যোগ হয়েছে!\n🔑 Keyword: `{keyword}`\n💬 উত্তর: {answer}",
        parse_mode="Markdown")


# ─────────────────────────────────────────────
#  Auto-FAQ message handler (user-side)
# ─────────────────────────────────────────────
@bot.message_handler(func=lambda m: (
    cfg("faq_enabled") is True
    and m.text
    and not m.text.startswith("/")
    and str(m.chat.id) != str(ADMIN_ID)
))
def auto_faq_reply(message):
    text  = message.text.lower().strip()
    items = cfg("faq_items") or []
    for item in items:
        if item.get("keyword", "").lower() in text:
            try:
                bot.send_message(message.chat.id, item["answer"], parse_mode="Markdown")
            except Exception:
                pass
            return


# ─────────────────────────────────────────────
#  Weekly Auto-Report
# ─────────────────────────────────────────────
def _send_weekly_report():
    total_users = len(user_data)
    total_dep   = sum(u.get("total_deposit", 0) for u in user_data.values())
    total_ord   = sum(u.get("total_orders", 0)  for u in user_data.values())
    total_bal   = sum(u.get("balance", 0)        for u in user_data.values())
    now_str     = str(bst_now())[:16]
    txt = (
        f"📊 *সাপ্তাহিক রিপোর্ট*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"📅 তারিখ: *{now_str}*\n\n"
        f"👥 মোট ইউজার: *{total_users}*\n"
        f"💰 মোট ডিপোজিট: *{total_dep} BDT*\n"
        f"🛒 মোট অর্ডার: *{total_ord}*\n"
        f"🏦 ইউজারদের মোট ব্যালেন্স: *{round(total_bal,2)} BDT*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨"
    )
    try:
        bot.send_message(ADMIN_ID, txt, parse_mode="Markdown")
    except Exception:
        pass
    settings["last_weekly_report"] = str(bst_now())
    save_settings(settings)


def _weekly_report_thread():
    import time as _time
    while True:
        _time.sleep(3600)
        try:
            last = cfg("last_weekly_report")
            if last:
                from datetime import timezone
                last_dt = datetime.fromisoformat(last).replace(tzinfo=timezone.utc) \
                          if last else None
                diff = (datetime.now(timezone.utc) - last_dt).total_seconds()
                if diff < 7 * 86400:
                    continue
            _send_weekly_report()
        except Exception:
            pass


# ─────────────────────────────────────────────
#  Stats
# ─────────────────────────────────────────────
def _adm_stats(call):
    total_users = len(user_data)
    total_dep   = sum(u.get("total_deposit", 0) for u in user_data.values())
    total_ord   = sum(u.get("total_orders", 0)  for u in user_data.values())
    total_bal   = sum(u.get("balance", 0)        for u in user_data.values())
    txt = (
        f"📊 *BOT STATISTICS*\n✨━━━━━━━━━━━━✨\n"
        f"👥 Total Users: *{total_users}*\n"
        f"💰 Total Deposits: *{total_dep} BDT*\n"
        f"🛒 Total Orders: *{total_ord}*\n"
        f"🏦 Users' Balance: *{round(total_bal,2)} BDT*\n"
        f"🤖 Userbot: *{userbot_status()}*\n"
        f"✨━━━━━━━━━━━━✨"
    )
    mk = types.InlineKeyboardMarkup()
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="adm|back"))
    try: bot.edit_message_text(txt, ADMIN_ID, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")


@bot.message_handler(commands=["stats"])
def cmd_stats(message):
    if not admin_only(message): return
    _adm_stats(type("FakeCall", (), {"message": message, "id": ""})())


@bot.message_handler(commands=["mybalance"])
def cmd_mybalance(message):
    """Admin command — fetch live supplier bot balance on demand."""
    if not admin_only(message): return
    if not _userbot or not cfg("supplier_bot_username"):
        bot.send_message(ADMIN_ID,
            "❌ Userbot সংযুক্ত নেই অথবা `supplier_bot_username` সেট করা হয়নি।",
            parse_mode="Markdown")
        return
    wait_msg = bot.send_message(ADMIN_ID, "⏳ সাপ্লায়ার বট থেকে ব্যালেন্স আনা হচ্ছে…")
    def _fetch():
        bal = _run_async(_check_supplier_balance_async(), timeout=40)
        try: bot.delete_message(ADMIN_ID, wait_msg.message_id)
        except Exception: pass
        if bal is None:
            bot.send_message(ADMIN_ID,
                "❌ *ব্যালেন্স পড়া যায়নি।*\n"
                "Userbot connected আছে কিনা এবং supplier bot username ঠিক আছে কিনা চেক করুন।",
                parse_mode="Markdown")
            return
        status = "✅ স্বাভাবিক" if bal >= _LOW_BAL_THRESHOLD else f"🔴 লো! ({_LOW_BAL_THRESHOLD} BDT-এর নিচে)"
        bot.send_message(ADMIN_ID,
            f"💰 *সাপ্লায়ার বট ব্যালেন্স*\n"
            f"✨━━━━━━━━━━━━━━━━━━✨\n"
            f"💎 ব্যালেন্স: *{bal} BDT*\n"
            f"📊 স্ট্যাটাস: {status}\n"
            f"✨━━━━━━━━━━━━━━━━━━✨",
            parse_mode="Markdown")
    threading.Thread(target=_fetch, daemon=True).start()


# ─────────────────────────────────────────────
#  Stock Sync (Userbot)
# ─────────────────────────────────────────────
def _adm_syncstock(call):
    supplier = cfg("supplier_bot_username")
    if not _userbot:
        bot.send_message(ADMIN_ID,
            "❌ *Userbot not connected!*\nSet `USER_API_ID`, `USER_API_HASH`, and `USER_SESSION_STRING` in Secrets.",
            parse_mode="Markdown"); return
    if not supplier:
        bot.send_message(ADMIN_ID,
            "❌ *Supplier Bot username not set!*\nGo to ⚙️ Global Config → 🤖 Supplier Bot.",
            parse_mode="Markdown"); return
    note = bot.send_message(ADMIN_ID,
        f"⏳ *Syncing stock from {supplier}...*\nSending `{cfg('supplier_stock_cmd')}` — please wait.",
        parse_mode="Markdown")
    result = _run_async(_sync_stock_async(), timeout=45)
    try: bot.delete_message(ADMIN_ID, note.message_id)
    except Exception: pass
    if result is None or result[0] is None:
        bot.send_message(ADMIN_ID,
            "❌ *Sync Failed!*\nSupplier bot did not respond within 40 seconds.\n"
            "Check `supplier_bot_username` and ensure the userbot is connected.",
            parse_mode="Markdown"); return
    updates, raw = result
    if not updates:
        bot.send_message(ADMIN_ID,
            f"⚠️ *No parseable stock data in reply.*\n\n*Raw reply:*\n```\n{raw[:400]}\n```",
            parse_mode="Markdown"); return
    # Update synced_stock in market_data
    prods   = get_products(); matched = 0
    report  = f"🔄 *Stock Sync Report*\n✨━━━━━━━━━━━━✨\n"
    for p_name in prods:
        # Fuzzy-match product name in updates
        for update_name, count in updates.items():
            if update_name.lower() in p_name.lower() or p_name.lower() in update_name.lower():
                market_data["products"][p_name]["synced_stock"] = count
                report += f"📦 *{p_name}*: `{count}` pcs\n"
                matched += 1
                break
    if matched: save_market_data(market_data)
    report += f"\n✅ *{matched} products synced!*\n✨━━━━━━━━━━━━✨"
    mk = types.InlineKeyboardMarkup()
    mk.add(types.InlineKeyboardButton("🔙 Admin Panel", callback_data="adm|back"))
    bot.send_message(ADMIN_ID, report, reply_markup=mk, parse_mode="Markdown")


# ─────────────────────────────────────────────────────────────────────
#  📊 Mail Report — Used (sold) vs Fresh (remaining) accounts as .xlsx
# ─────────────────────────────────────────────────────────────────────
@bot.message_handler(commands=["mailstock", "stocklist"])
def admin_mail_stock(message):
    """Quick command: instantly shows Fresh vs Used mail stock for every product."""
    if message.chat.id != ADMIN_ID: return
    prods = {p: info for p, info in get_products().items() if info.get("cat") != "vpn"}
    if not prods:
        bot.send_message(ADMIN_ID, "❌ কোনো মেইল প্রোডাক্ট পাওয়া যায়নি।"); return
    lines = ["📊 *Mail Stock Overview*", "✨━━━━━━━━━━━━━━━━━━✨"]
    total_fresh = total_used = 0
    for p in prods:
        fresh = get_stock_count(p)
        used  = len(db_load_sold(p))
        total_fresh += fresh
        total_used  += used
        warn = " ⚠️" if fresh <= 20 else ""
        lines.append(f"📧 *{p}*\n   🟢 Fresh: *{fresh}*  |  ✅ Used: *{used}*{warn}")
    lines.append("✨━━━━━━━━━━━━━━━━━━✨")
    lines.append(f"📦 *Total Fresh:* {total_fresh}  |  *Total Used:* {total_used}")
    mk = types.InlineKeyboardMarkup(row_width=1)
    mk.add(types.InlineKeyboardButton("📥 Download Detailed Report", callback_data="adm|mailreport"))
    bot.send_message(ADMIN_ID, "\n".join(lines), reply_markup=mk, parse_mode="Markdown")


def _adm_mailreport_select(call):
    """List mail products (non-VPN) so admin can pick which one to report on."""
    prods = {p: info for p, info in get_products().items() if info.get("cat") != "vpn"}
    if not prods:
        bot.send_message(ADMIN_ID, "❌ কোনো মেইল প্রোডাক্ট পাওয়া যায়নি।"); return
    mk = types.InlineKeyboardMarkup(row_width=1)
    for p in prods:
        fresh = get_stock_count(p)
        used  = len(db_load_sold(p))
        mk.add(types.InlineKeyboardButton(
            f"📧 {p}  (Fresh: {fresh} | Used: {used})", callback_data=f"mrep|{p}"))
    mk.add(types.InlineKeyboardButton("🔙 Admin Panel", callback_data="adm|back"))
    bot.send_message(ADMIN_ID, "📊 কোন প্রোডাক্টের রিপোর্ট দেখতে চান?", reply_markup=mk)


@bot.callback_query_handler(func=lambda c: c.data.startswith("mrep|"))
def _adm_mailreport_kind(call):
    if call.message.chat.id != ADMIN_ID: return
    p_name = call.data.split("|", 1)[1]
    fresh  = get_stock_count(p_name)
    used   = len(db_load_sold(p_name))
    mk = types.InlineKeyboardMarkup(row_width=1)
    mk.add(types.InlineKeyboardButton(f"📥 Fresh Mails ({fresh}) — .xlsx", callback_data=f"mrepf|fresh|{p_name}"))
    mk.add(types.InlineKeyboardButton(f"✅ Used Mails ({used}) — .xlsx",  callback_data=f"mrepf|used|{p_name}"))
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="adm|mailreport"))
    bot.answer_callback_query(call.id)
    bot.send_message(ADMIN_ID, f"📧 *{p_name}*\nFresh: *{fresh}* | Used: *{used}*",
                     reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("mrepf|"))
def _adm_mailreport_send(call):
    if call.message.chat.id != ADMIN_ID: return
    _, kind, p_name = call.data.split("|", 2)
    rows = _load_stock_rows(p_name) if kind == "fresh" else db_load_sold(p_name)
    bot.answer_callback_query(call.id)
    if not rows:
        bot.send_message(ADMIN_ID, f"❌ *{p_name}* — {'ফ্রেশ' if kind=='fresh' else 'ইউজড'} মেইল নেই।",
                         parse_mode="Markdown"); return
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, index=False)
    out.seek(0)
    label = "Fresh" if kind == "fresh" else "Used"
    fname = f"{p_name.replace(' ', '_')}_{label}_{len(rows)}.xlsx"
    bot.send_document(ADMIN_ID, out, visible_file_name=fname,
                      caption=f"📊 *{p_name}* — {label} mails: *{len(rows)}* পিস",
                      parse_mode="Markdown")


# ─────────────────────────────────────────────────────────────────────
#  🤖 Supplier Settings Admin Panel
# ─────────────────────────────────────────────────────────────────────
def _adm_userbot(call):
    status    = userbot_status()
    supplier  = cfg("supplier_bot_username")         or "(not set)"
    init_cmd  = cfg("supplier_initial_cmd")          or "/start"
    vpn_btn   = cfg("supplier_vpn_button")           or "🛡️ Buy VPN"
    dur_btns  = cfg("supplier_duration_buttons")     or {}
    confirm   = cfg("supplier_confirm_button")       or "✅ Confirm"
    stk_cmd   = cfg("supplier_stock_cmd")            or "/stock"

    dur_lines = "\n".join(f"  `{k}` → `{v}`" for k, v in dur_btns.items()) or "  (not set)"
    txt = (
        "🤖 *Supplier Bot Settings*\n"
        "✨━━━━━━━━━━━━━━━━━━✨\n"
        f"🔌 Userbot Status: *{status}*\n\n"
        f"🏪 *Supplier Bot:* `{supplier}`\n"
        f"▶️ *Initial Command:* `{init_cmd}`\n"
        f"🛡️ *VPN Button Name:* `{vpn_btn}`\n"
        f"⏱️ *Duration Buttons:*\n{dur_lines}\n"
        f"✅ *Confirm Button:* `{confirm}`\n"
        f"📦 *Stock Command:* `{stk_cmd}`\n"
        "✨━━━━━━━━━━━━━━━━━━✨\n"
        "_Credentials load from Replit Secrets automatically._"
    )
    ub_on = cfg("userbot_enabled") is not False
    mk = types.InlineKeyboardMarkup(row_width=2)
    mk.add(
        types.InlineKeyboardButton(
            f"{'✅ Userbot ON' if ub_on else '❌ Userbot OFF'} — Toggle",
            callback_data="toggle|userbot_enabled"
        ),
    )
    mk.add(
        types.InlineKeyboardButton("🏪 Supplier Username", callback_data="ub|supplier"),
        types.InlineKeyboardButton("▶️ Initial Command",   callback_data="ub|init_cmd"),
    )
    mk.add(
        types.InlineKeyboardButton("🛡️ VPN Button Name",   callback_data="ub|vpn_btn"),
        types.InlineKeyboardButton("⏱️ Duration Buttons",  callback_data="ub|dur_btns"),
    )
    mk.add(
        types.InlineKeyboardButton("✅ Confirm Button",    callback_data="ub|confirm_btn"),
        types.InlineKeyboardButton("📦 Stock Command",     callback_data="ub|stock_cmd"),
    )
    mk.add(
        types.InlineKeyboardButton("🔗 Map Buttons",       callback_data="ub|map_btns"),
        types.InlineKeyboardButton("🔄 Sync Stock Now",    callback_data="adm|syncstock"),
    )
    mk.add(
        types.InlineKeyboardButton("🔙 Back",              callback_data="adm|back"),
    )
    try:
        bot.edit_message_text(txt, ADMIN_ID, call.message.message_id,
                              reply_markup=mk, parse_mode="Markdown")
    except Exception:
        bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")


def _adm_userbot_dur_btns(call):
    """Sub-menu: edit per-duration button names shown in the supplier bot."""
    dur_btns = cfg("supplier_duration_buttons") or {}
    durs     = get_vpn_durations()
    mk = types.InlineKeyboardMarkup(row_width=2)
    for dk, dv in durs.items():
        current = dur_btns.get(dk, "(not set)")
        mk.add(types.InlineKeyboardButton(
            f"✏️ {dv['name']}", callback_data=f"ub_dur|{dk}"))
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="ub|back_supplier"))
    txt = (
        "⏱️ *Duration Button Names*\n"
        "✨━━━━━━━━━━━━✨\n"
        "Set the *exact* button text shown in the supplier bot for each duration.\n\n"
    )
    for dk, dv in durs.items():
        current = dur_btns.get(dk, "_(not set)_")
        txt += f"  • {dv['name']} → `{current}`\n"
    txt += "\n✨━━━━━━━━━━━━✨\n_Tap a row to update._"
    try:
        bot.edit_message_text(txt, ADMIN_ID, call.message.message_id,
                              reply_markup=mk, parse_mode="Markdown")
    except Exception:
        bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("ub_dur|"))
def ub_dur_edit_start(call):
    if call.message.chat.id != ADMIN_ID: return
    dk = call.data.split("|")[1]; bot.answer_callback_query(call.id)
    durs     = get_vpn_durations()
    dur_btns = cfg("supplier_duration_buttons") or {}
    current  = dur_btns.get(dk, "(not set)")
    dur_name = durs.get(dk, {}).get("name", dk)
    msg = bot.send_message(
        ADMIN_ID,
        f"⏱️ *{dur_name}*\n"
        f"Current supplier button text: `{current}`\n\n"
        "Send the *exact* button text shown in the supplier bot\n"
        "(e.g. `3 Days VPN` or `07 Days`):",
        parse_mode="Markdown"
    )
    bot.register_next_step_handler(msg, _ub_save_dur_btn, dk)


def _ub_save_dur_btn(message, dk):
    if message.chat.id != ADMIN_ID: return
    val      = message.text.strip()
    dur_btns = cfg("supplier_duration_buttons") or {}
    dur_btns[dk] = val
    settings["supplier_duration_buttons"] = dur_btns
    save_settings(settings)
    durs     = get_vpn_durations()
    dur_name = durs.get(dk, {}).get("name", dk)
    bot.send_message(ADMIN_ID,
        f"✅ *{dur_name}* button set to:\n`{val}`", parse_mode="Markdown")


def _adm_map_buttons(call):
    """Sub-menu: view and edit the button_map (my name → supplier button name)."""
    btn_map  = cfg("button_map") or {}
    durs     = get_vpn_durations()
    prods    = get_products()
    mk = types.InlineKeyboardMarkup(row_width=2)

    # Duration mappings
    for dk, dv in durs.items():
        raw   = (cfg("supplier_duration_buttons") or {}).get(dk, dk)
        label = btn_map.get(raw, "_(not set)_")
        mk.add(types.InlineKeyboardButton(
            f"⏱️ {dv['name']}", callback_data=f"ub_map|dur|{dk}"))

    # Product mappings (VPN only)
    for p_name, p_info in prods.items():
        if p_info.get("cat") == "vpn":
            mk.add(types.InlineKeyboardButton(
                f"🛡️ {p_name}", callback_data=f"ub_map|prod|{p_name}"))

    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="ub|back_supplier"))

    lines = []
    for dk, dv in durs.items():
        raw = (cfg("supplier_duration_buttons") or {}).get(dk, dk)
        mapped = btn_map.get(raw, "_(not set)_")
        lines.append(f"  ⏱️ `{raw}` → `{mapped}`")
    for p_name, p_info in prods.items():
        if p_info.get("cat") == "vpn":
            mapped = btn_map.get(p_name, "_(not set)_")
            lines.append(f"  🛡️ `{p_name}` → `{mapped}`")

    txt = (
        "🔗 *Button Mapping*\n"
        "✨━━━━━━━━━━━━✨\n"
        "Map your names to the *exact* supplier bot button text.\n\n"
        + "\n".join(lines) +
        "\n\n✨━━━━━━━━━━━━✨\n_Tap a row to set or update its mapping._"
    )
    try:
        bot.edit_message_text(txt, ADMIN_ID, call.message.message_id,
                              reply_markup=mk, parse_mode="Markdown")
    except Exception:
        bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("ub_map|"))
def ub_map_edit_start(call):
    if call.message.chat.id != ADMIN_ID: return
    parts  = call.data.split("|")       # ub_map | dur/prod | key
    kind   = parts[1]
    key    = parts[2]
    btn_map = cfg("button_map") or {}
    bot.answer_callback_query(call.id)

    if kind == "dur":
        durs    = get_vpn_durations()
        raw     = (cfg("supplier_duration_buttons") or {}).get(key, key)
        label   = durs.get(key, {}).get("name", key)
        current = btn_map.get(raw, "(not set)")
        prompt  = (
            f"🔗 *Map Duration: {label}*\n"
            f"My duration key: `{raw}`\n"
            f"Current supplier button: `{current}`\n\n"
            "Send the *exact* button text in the supplier bot\n"
            "(e.g. `3 Days VPN` or `07 Days`):"
        )
        map_key = raw
    else:
        current = btn_map.get(key, "(not set)")
        prompt  = (
            f"🔗 *Map Product: {key}*\n"
            f"Current supplier button: `{current}`\n\n"
            "Send the *exact* button text in the supplier bot\n"
            "(e.g. `Nord-Premium` or `Express VPN`):"
        )
        map_key = key

    msg = bot.send_message(ADMIN_ID, prompt, parse_mode="Markdown")
    bot.register_next_step_handler(msg, _ub_save_map, map_key)


def _ub_save_map(message, map_key):
    if message.chat.id != ADMIN_ID: return
    val = message.text.strip()
    if not val:
        bot.send_message(ADMIN_ID, "❌ Empty input. Mapping not saved.")
        return
    btn_map = cfg("button_map") or {}
    btn_map[map_key] = val
    settings["button_map"] = btn_map
    save_settings(settings)
    bot.send_message(ADMIN_ID,
        f"✅ *Mapped:*\n`{map_key}` → `{val}`\n\n"
        "The userbot will now search for this button name in the supplier bot.",
        parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("ub|"))
def ub_router(call):
    if call.message.chat.id != ADMIN_ID: return
    key = call.data.split("|")[1]; bot.answer_callback_query(call.id)

    if key == "back_supplier":
        _adm_userbot(call); return

    if key == "dur_btns":
        _adm_userbot_dur_btns(call); return

    if key == "map_btns":
        _adm_map_buttons(call); return

    prompts = {
        "supplier":    (
            "🏪 *Supplier Bot Username*\nE.g. `@PremiumShopbd_bot`:",
            "supplier_bot_username"),
        "init_cmd":    (
            "▶️ *Initial Command*\nSent to start a purchase session.\nE.g. `/start` or `/buy`:",
            "supplier_initial_cmd"),
        "vpn_btn":     (
            "🛡️ *VPN Category Button Name*\nExact text of the VPN category button.\nE.g. `🛡️ Buy VPN`:",
            "supplier_vpn_button"),
        "confirm_btn": (
            "✅ *Confirm Button Name*\nExact text of the final confirm button.\nE.g. `✅ Confirm`:",
            "supplier_confirm_button"),
        "stock_cmd":   (
            "📦 *Stock Command*\nE.g. `/stock`:",
            "supplier_stock_cmd"),
        "buy_cmd":     (
            "🛒 *Buy Command (legacy fallback)*\nUse `{product}` and `{qty}` as placeholders:",
            "supplier_buy_cmd"),
    }
    if key in prompts:
        prompt, setting_key = prompts[key]
        msg = bot.send_message(ADMIN_ID, prompt, parse_mode="Markdown")
        bot.register_next_step_handler(msg, _ub_save_str, setting_key)


def _ub_save_str(message, key):
    if message.chat.id != ADMIN_ID: return
    val = message.text.strip()
    settings[key] = val
    save_settings(settings)
    bot.send_message(ADMIN_ID, f"✅ *{key}* updated to:\n`{val}`", parse_mode="Markdown")


# ─────────────────────────────────────────────
#  User Manager
# ─────────────────────────────────────────────
def _adm_usermgr(call):
    total = len(user_data)
    pending_cnt  = len(pending_deposits)
    rejected_cnt = len(rejected_deposits)
    mk = types.InlineKeyboardMarkup(row_width=2)
    mk.add(
        types.InlineKeyboardButton("🔍 Search by ID",       callback_data="usrmgr|search"),
        types.InlineKeyboardButton("👤 Search by Username", callback_data="usrmgr|searchname"),
    )
    mk.add(
        types.InlineKeyboardButton("📋 Users List",         callback_data="usrmgr|list"),
        types.InlineKeyboardButton("🚫 Banned Users",       callback_data="usrmgr|banned"),
    )
    mk.add(
        types.InlineKeyboardButton(f"💳 Pending Deposits ({pending_cnt})",  callback_data="usrmgr|pending"),
        types.InlineKeyboardButton(f"❌ Rejected Deposits ({rejected_cnt})", callback_data="usrmgr|rejected"),
    )
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="adm|back"))
    txt = (
        f"👥 *User Manager*\n✨━━━━━━━━━━━━✨\n"
        f"👤 Total Users: *{total}*\n"
        f"💳 Pending: *{pending_cnt}* | ❌ Rejected: *{rejected_cnt}*\n"
        "✨━━━━━━━━━━━━✨"
    )
    try:
        bot.edit_message_text(txt, ADMIN_ID, call.message.message_id,
                              reply_markup=mk, parse_mode="Markdown")
    except Exception:
        bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("usrmgr|"))
def usrmgr_router(call):
    if call.message.chat.id != ADMIN_ID: return
    action = call.data.split("|")[1]; bot.answer_callback_query(call.id)
    if action == "search":
        msg = bot.send_message(ADMIN_ID, "🔍 ইউজারের *ID* দিন (যেমন: `7522357347`):", parse_mode="Markdown")
        bot.register_next_step_handler(msg, _usrmgr_show_by_input)
    elif action == "searchname":
        msg = bot.send_message(ADMIN_ID, "👤 ইউজারের *Username* দিন (যেমন: `@johndoe` বা `johndoe`):", parse_mode="Markdown")
        bot.register_next_step_handler(msg, _usrmgr_show_by_username)
    elif action == "list":
        # Show sub-options for how many users to list
        mk2 = types.InlineKeyboardMarkup(row_width=1)
        mk2.add(types.InlineKeyboardButton("📋 Last 10 Users",    callback_data="usrmgr|list10"))
        mk2.add(types.InlineKeyboardButton("📋 Last 20 Users",    callback_data="usrmgr|list20"))
        mk2.add(types.InlineKeyboardButton("📋 All Users List",   callback_data="usrmgr|listall"))
        bot.send_message(ADMIN_ID, "📋 *কতজন ইউজার দেখতে চান?*", reply_markup=mk2, parse_mode="Markdown")
    elif action in ("list10", "list20", "listall"):
        uids = list(user_data.keys())
        if action == "list10":   subset = uids[-10:];  title = "Last 10"
        elif action == "list20": subset = uids[-20:];  title = "Last 20"
        else:                    subset = uids;        title = f"All {len(uids)}"
        # Build XLSX with full user details
        rows = []
        for uid in subset:
            u = user_data[uid]
            raw_un = u.get("username", "") or ""
            username = ("@" + raw_un.lstrip("@")) if raw_un.strip() else "—"
            rows.append({
                "Telegram ID":       uid,
                "Username":          username,
                "Name":              u.get("first_name", "") or "—",
                "Balance (BDT)":     u.get("balance", 0),
                "Total Deposit (BDT)": u.get("total_deposit", 0),
                "Total Orders":      u.get("total_orders", 0),
                "Status":            "🚫 Banned" if u.get("banned") else "✅ Active",
                "Language":          u.get("language", "bn"),
                "Last Deposit":      u.get("last_deposit_time") or "—",
                "Last Purchase":     u.get("last_purchase_time") or "—",
            })
        df = pd.DataFrame(rows)
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Users")
            ws = writer.sheets["Users"]
            # Auto-fit column widths
            for col in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)
        out.seek(0)
        fname = f"users_{title.replace(' ', '_').lower()}_{len(rows)}.xlsx"
        bot.send_document(
            ADMIN_ID, out, visible_file_name=fname,
            caption=f"📊 *{title} Users* — {len(rows)} জন\n📁 ফাইলে সব ডিটেইলস আছে।",
            parse_mode="Markdown"
        )
    elif action == "banned":
        banned = [uid for uid, u in user_data.items() if u.get("banned")]
        if not banned: bot.send_message(ADMIN_ID, "✅ No banned users."); return
        bot.send_message(ADMIN_ID, "🚫 *Banned Users:*\n" + "\n".join(f"`{uid}`" for uid in banned), parse_mode="Markdown")
    elif action == "pending":
        if not pending_deposits:
            bot.send_message(ADMIN_ID, "✅ *কোনো Pending Deposit নেই।*", parse_mode="Markdown"); return
        bot.send_message(ADMIN_ID,
            f"💳 *Pending Deposits ({len(pending_deposits)})*\n✨━━━━━━━━━━━━✨",
            parse_mode="Markdown")
        for uid, d in list(pending_deposits.items()):
            amt_txt = f"${d.get('usd','')} = {d['bdt']} BDT" if d.get("usd") else f"{d['bdt']} BDT"
            txt = (f"👤 `{uid}`\n💳 {d['method']} — {amt_txt}\n"
                   f"📋 TrxID: `{d['trx']}`\n🕐 {d.get('time','')}")
            mk2 = types.InlineKeyboardMarkup(row_width=2)
            trx_k = (d.get("trx_key") or d["trx"].lower()[:20])
            mk2.add(
                types.InlineKeyboardButton("✅ Approve", callback_data=f"adm_aprv_{uid}_{int(d['bdt'])}_{trx_k[:20]}"),
                types.InlineKeyboardButton("❌ Reject",  callback_data=f"adm_rej_{uid}"),
            )
            bot.send_message(ADMIN_ID, txt, reply_markup=mk2, parse_mode="Markdown")
    elif action == "rejected":
        if not rejected_deposits:
            bot.send_message(ADMIN_ID, "✅ *কোনো Rejected Deposit নেই।*", parse_mode="Markdown"); return
        recent = rejected_deposits[-30:]
        lines = []
        for d in reversed(recent):
            amt_txt = f"${d.get('usd','')} = {d['bdt']} BDT" if d.get("usd") else f"{d['bdt']} BDT"
            lines.append(f"👤 `{d['uid']}` — {d['method']} {amt_txt}\n"
                         f"   TrxID: `{d['trx']}` | 🕐 {d.get('rejected_at','')}")
        txt = f"❌ *Rejected Deposits (Last 30)*\n✨━━━━━━━━━━━━✨\n\n" + "\n\n".join(lines)
        # Split if too long
        if len(txt) > 4000:
            for i in range(0, len(lines), 10):
                chunk_txt = "\n\n".join(lines[i:i+10])
                bot.send_message(ADMIN_ID, chunk_txt, parse_mode="Markdown")
        else:
            bot.send_message(ADMIN_ID, txt, parse_mode="Markdown")


def _usrmgr_show_by_input(message):
    if message.chat.id != ADMIN_ID: return
    uid = message.text.strip()
    if uid not in user_data:
        bot.send_message(ADMIN_ID, "❌ *User not found.*\nID দিয়ে কোনো ইউজার পাওয়া যায়নি।", parse_mode="Markdown"); return
    _send_user_detail(uid)


def _usrmgr_show_by_username(message):
    if message.chat.id != ADMIN_ID: return
    query = message.text.strip().lstrip("@").lower()
    matches = []
    for uid, u in user_data.items():
        stored_un = u.get("username", "").lstrip("@").lower()
        stored_fn = u.get("first_name", "").lower()
        if query and (query == stored_un or query in stored_un or query in stored_fn):
            matches.append(uid)
    if not matches:
        bot.send_message(ADMIN_ID,
            f"❌ *@{query}* নামে কোনো ইউজার পাওয়া যায়নি।\n"
            "_(ইউজারকে আগে বটে মেসেজ পাঠাতে হবে যাতে username সেভ হয়)_",
            parse_mode="Markdown"); return
    if len(matches) == 1:
        _send_user_detail(matches[0]); return
    lines = []
    for uid in matches[:20]:
        u = user_data[uid]
        lines.append(f"`{uid}` — {u.get('username','')} {u.get('first_name','')} | 💰{u.get('balance',0)} BDT")
    bot.send_message(ADMIN_ID,
        f"👤 *{len(matches)} জন ইউজার পাওয়া গেছে:*\n✨━━━━━━━━━━━━✨\n" + "\n".join(lines),
        parse_mode="Markdown")


def _send_user_detail(uid):
    u = user_data.get(uid, {}); ban_status = "🚫 Banned" if u.get("banned") else "✅ Active"
    txt = (
        f"👤 *USER DETAILS*\n✨━━━━━━━━━━━━✨\n"
        f"🆔 ID: `{uid}`\n🌐 Language: *{u.get('language','bn')}*\n🔰 Status: {ban_status}\n"
        f"✨━━━━━━━━━━━━✨\n"
        f"💰 Balance: *{u.get('balance',0)} BDT*\n"
        f"📈 Total Deposit: *{u.get('total_deposit',0)} BDT*\n"
        f"🛒 Total Orders: *{u.get('total_orders',0)}*\n"
        f"✨━━━━━━━━━━━━✨\n"
        f"⏳ Last Deposit: {u.get('last_deposit_time','N/A')}\n"
        f"🎁 Last Purchase: {u.get('last_purchase_time','N/A')}\n✨━━━━━━━━━━━━✨"
    )
    mk = types.InlineKeyboardMarkup(row_width=2)
    mk.add(
        types.InlineKeyboardButton("💰 Edit Balance", callback_data=f"musr_bal|{uid}"),
        types.InlineKeyboardButton("🛒 Edit Orders",  callback_data=f"musr_ord|{uid}"),
    )
    mk.add(types.InlineKeyboardButton("📩 Send Message", callback_data=f"musr_msg|{uid}"))
    if u.get("banned"):
        mk.add(types.InlineKeyboardButton("✅ Unban User", callback_data=f"musr_unban|{uid}"))
    else:
        mk.add(types.InlineKeyboardButton("🚫 Ban User",   callback_data=f"musr_ban|{uid}"))
    mk.add(types.InlineKeyboardButton("🗑️ Delete User", callback_data=f"musr_del|{uid}"))
    bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("musr_"))
def manage_user_action(call):
    if call.message.chat.id != ADMIN_ID: return
    parts = call.data.split("|", 1); action = parts[0]; uid = parts[1]
    if action == "musr_bal":
        bot.answer_callback_query(call.id)
        msg = bot.send_message(ADMIN_ID,
            f"💰 User `{uid}`\nCurrent: *{user_data.get(uid,{}).get('balance',0)} BDT*\n\nNew balance:",
            parse_mode="Markdown")
        bot.register_next_step_handler(msg, _musr_set_balance, uid)
    elif action == "musr_ord":
        bot.answer_callback_query(call.id)
        msg = bot.send_message(ADMIN_ID,
            f"🛒 User `{uid}`\nCurrent: *{user_data.get(uid,{}).get('total_orders',0)}*\n\nNew count:",
            parse_mode="Markdown")
        bot.register_next_step_handler(msg, _musr_set_orders, uid)
    elif action == "musr_ban":
        if uid in user_data:
            user_data[uid]["banned"] = True; save_data(user_data)
            bot.answer_callback_query(call.id, "🚫 Banned!", show_alert=True)
            try: bot.send_message(uid, "🚫 আপনাকে এই বট থেকে ব্যান করা হয়েছে।")
            except Exception: pass
            _send_user_detail(uid)
    elif action == "musr_unban":
        if uid in user_data:
            user_data[uid]["banned"] = False; save_data(user_data)
            bot.answer_callback_query(call.id, "✅ Unbanned!", show_alert=True)
            try: bot.send_message(uid, "✅ আপনার ব্যান তুলে নেওয়া হয়েছে।")
            except Exception: pass
            _send_user_detail(uid)
    elif action == "musr_msg":
        bot.answer_callback_query(call.id)
        msg = bot.send_message(ADMIN_ID,
            f"📩 User `{uid}` কে কী পাঠাতে চান?\n_(নোটিফিকেশন, ওয়ার্নিং, যেকোনো মেসেজ)_",
            parse_mode="Markdown")
        bot.register_next_step_handler(msg, _musr_send_message, uid)
    elif action == "musr_del":
        if uid in user_data:
            del user_data[uid]; save_data(user_data)
            bot.answer_callback_query(call.id, "🗑️ Deleted!", show_alert=True)
            bot.edit_message_text(f"🗑️ User `{uid}` deleted.",
                                  ADMIN_ID, call.message.message_id, parse_mode="Markdown")


def _musr_set_balance(message, uid):
    if message.chat.id != ADMIN_ID: return
    try:
        new_bal = round(float(message.text.strip()), 2)
        user_data[uid]["balance"] = new_bal; save_data(user_data)
        bot.send_message(ADMIN_ID, f"✅ Balance → *{new_bal} BDT* for `{uid}`", parse_mode="Markdown")
        try: bot.send_message(uid, f"💰 আপনার ব্যালেন্স আপডেট: *{new_bal} BDT*", parse_mode="Markdown")
        except Exception: pass
        _send_user_detail(uid)
    except Exception: bot.send_message(ADMIN_ID, "❌ Invalid. Enter a number.")


def _musr_set_orders(message, uid):
    if message.chat.id != ADMIN_ID: return
    try:
        new_count = int(message.text.strip())
        user_data[uid]["total_orders"] = new_count; save_data(user_data)
        bot.send_message(ADMIN_ID, f"✅ Orders → *{new_count}* for `{uid}`", parse_mode="Markdown")
        _send_user_detail(uid)
    except Exception: bot.send_message(ADMIN_ID, "❌ Invalid. Enter an integer.")


def _musr_send_message(message, uid):
    if message.chat.id != ADMIN_ID: return
    text = message.text.strip() if message.text else ""
    if not text:
        bot.send_message(ADMIN_ID, "❌ খালি মেসেজ পাঠানো যাবে না।"); return
    try:
        bot.send_message(uid,
            f"📣 *Admin নোটিফিকেশন:*\n"
            f"✨━━━━━━━━━━━━━━━━━━✨\n"
            f"{text}\n"
            f"✨━━━━━━━━━━━━━━━━━━✨",
            parse_mode="Markdown")
        bot.send_message(ADMIN_ID, f"✅ মেসেজ পাঠানো হয়েছে → `{uid}`", parse_mode="Markdown")
    except Exception as e:
        bot.send_message(ADMIN_ID, f"❌ পাঠাতে ব্যর্থ: {e}")


@bot.message_handler(commands=["manage_user"])
def manage_user_start(message):
    if not admin_only(message): return
    msg = bot.send_message(ADMIN_ID, "🆔 ইউজারের User ID লিখুন:")
    bot.register_next_step_handler(msg, _usrmgr_show_by_input)


# ─────────────────────────────────────────────
#  Branding Manager
# ─────────────────────────────────────────────
def _adm_branding(call):
    photo_url = cfg("welcome_photo_url") or "*(not set)*"
    welcome_t = cfg("welcome_text")      or "*(uses texts.json default)*"
    mk = types.InlineKeyboardMarkup(row_width=2)
    mk.add(
        types.InlineKeyboardButton("🖼️ Welcome Image URL", callback_data="brand|photo"),
        types.InlineKeyboardButton("📝 Welcome Text",      callback_data="brand|text"),
    )
    mk.add(
        types.InlineKeyboardButton("🗑️ Clear Image URL",   callback_data="brand|clearphoto"),
        types.InlineKeyboardButton("🔙 Back",              callback_data="adm|back"),
    )
    txt = (
        f"🎨 *Branding Manager*\n✨━━━━━━━━━━━━✨\n"
        f"🖼️ *Welcome Photo:*\n`{str(photo_url)[:80]}`\n\n"
        f"📝 *Welcome Text:*\n{str(welcome_t)[:100]}\n✨━━━━━━━━━━━━✨"
    )
    try: bot.edit_message_text(txt, ADMIN_ID, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("brand|"))
def branding_router(call):
    if call.message.chat.id != ADMIN_ID: return
    action = call.data.split("|")[1]; bot.answer_callback_query(call.id)
    if action == "photo":
        msg = bot.send_message(ADMIN_ID,
            "🖼️ Send new *Welcome Image URL*:\n_(Direct image link, e.g. https://i.imgur.com/x.jpg)_",
            parse_mode="Markdown")
        bot.register_next_step_handler(msg, _brand_set_photo)
    elif action == "text":
        msg = bot.send_message(ADMIN_ID,
            "📝 Send new *Welcome Text* (Markdown supported):",
            parse_mode="Markdown")
        bot.register_next_step_handler(msg, _brand_set_text)
    elif action == "clearphoto":
        settings["welcome_photo_url"] = ""; save_settings(settings)
        bot.send_message(ADMIN_ID, "✅ Welcome image URL cleared.")

def _brand_set_photo(message):
    if message.chat.id != ADMIN_ID: return
    url = message.text.strip(); settings["welcome_photo_url"] = url; save_settings(settings)
    bot.send_message(ADMIN_ID, f"✅ Welcome image URL updated!\n`{url}`", parse_mode="Markdown")

def _brand_set_text(message):
    if message.chat.id != ADMIN_ID: return
    text = message.text.strip(); settings["welcome_text"] = text; save_settings(settings)
    bot.send_message(ADMIN_ID, f"✅ Welcome text updated!", parse_mode="Markdown")


# ─────────────────────────────────────────────
#  Button Labels Manager
# ─────────────────────────────────────────────
_BTN_LABEL_KEYS = [
    ("btn_profile", "👤 Profile বাটন"),
    ("btn_deposit", "💰 Deposit বাটন"),
    ("btn_shop",    "🛒 Shop Now বাটন"),
    ("btn_price",   "💎 Price List বাটন"),
    ("btn_support", "☎️ Support বাটন"),
    ("btn_daily",   "🎁 Daily Bonus বাটন"),
]

def _adm_btnlabels(call):
    mk = types.InlineKeyboardMarkup(row_width=1)
    for key, label in _BTN_LABEL_KEYS:
        current = cfg(key) or "(default)"
        mk.add(types.InlineKeyboardButton(
            f"✏️ {label}  →  {current}", callback_data=f"btnlbl|{key}"))
    mk.add(types.InlineKeyboardButton("🔄 Reset All to Default", callback_data="btnlbl|reset_all"))
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="adm|back"))
    txt = (
        "🔘 *Button Labels Manager*\n"
        "✨━━━━━━━━━━━━━━━━━━✨\n"
        "নিচের যেকোনো বাটনে ক্লিক করুন এবং নতুন টেক্সট/ইমোজি লিখুন।\n\n"
        "💡 *টিপস:*\n"
        "🟣 = Violet ইফেক্ট\n"
        "🟢 = LawnGreen ইফেক্ট\n"
        "✨━━━━━━━━━━━━━━━━━━✨"
    )
    try: bot.edit_message_text(txt, ADMIN_ID, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("btnlbl|"))
def btnlbl_router(call):
    if call.message.chat.id != ADMIN_ID: return
    bot.answer_callback_query(call.id)
    key = call.data.split("|")[1]

    if key == "reset_all":
        defaults = {
            "btn_profile": "🟣 👤 Profile",
            "btn_deposit": "🟣 💰 Deposit",
            "btn_shop":    "🟣 🛒 Shop Now",
            "btn_price":   "🟣 💎 Price List",
            "btn_support": "🟢 ☎️ Support",
            "btn_daily":   "🎁 Daily Bonus",
        }
        for k, v in defaults.items():
            settings[k] = v
        save_settings(settings)
        bot.send_message(ADMIN_ID, "✅ সব বাটন লেবেল ডিফল্টে ফিরে গেছে!")
        _adm_btnlabels(call); return

    label_name = next((l for k, l in _BTN_LABEL_KEYS if k == key), key)
    current    = cfg(key) or "(default)"
    msg = bot.send_message(ADMIN_ID,
        f"✏️ *{label_name}*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"বর্তমান: `{current}`\n\n"
        f"নতুন টেক্সট লিখুন (emoji + text):\n"
        f"_উদাহরণ: 🟣 💰 জমা করুন_\n"
        f"✨━━━━━━━━━━━━━━━━━━✨",
        parse_mode="Markdown")
    bot.register_next_step_handler(msg, lambda m, k=key: _btnlbl_save(m, k))


def _btnlbl_save(message, key):
    if message.chat.id != ADMIN_ID: return
    new_label = message.text.strip()
    if not new_label:
        bot.send_message(ADMIN_ID, "❌ খালি লেবেল সেট করা যাবে না।"); return
    settings[key] = new_label
    save_settings(settings)
    label_name = next((l for k, l in _BTN_LABEL_KEYS if k == key), key)
    bot.send_message(ADMIN_ID,
        f"✅ *{label_name}* আপডেট হয়েছে!\n"
        f"নতুন লেবেল: `{new_label}`\n\n"
        f"ইউজাররা পরবর্তী /start থেকে নতুন বাটন দেখবে।",
        parse_mode="Markdown")


# ─────────────────────────────────────────────
#  Global Config
# ─────────────────────────────────────────────
def _adm_config(call):
    supplier = cfg("supplier_bot_username") or "(not set)"
    mk = types.InlineKeyboardMarkup(row_width=2)
    mk.add(
        types.InlineKeyboardButton("💵 USD Rate",         callback_data="cfg|usd_rate"),
        types.InlineKeyboardButton("👤 Support Username", callback_data="cfg|support_username"),
    )
    mk.add(
        types.InlineKeyboardButton("🎁 Daily Bonus Amt",  callback_data="cfg|daily_bonus_amount"),
        types.InlineKeyboardButton("🔧 Maintenance Msg",  callback_data="cfg|maintenance_message"),
    )
    mk.add(
        types.InlineKeyboardButton("🔗 Channel Links",    callback_data="cfg|channels"),
        types.InlineKeyboardButton("💳 Payment Methods",  callback_data="cfg|payments"),
    )
    mk.add(
        types.InlineKeyboardButton("🤖 Supplier Bot",     callback_data="cfg|supplier_bot"),
        types.InlineKeyboardButton("🔙 Back",             callback_data="adm|back"),
    )
    txt = (
        "⚙️ *Global Configuration*\n✨━━━━━━━━━━━━✨\n"
        f"💵 USD Rate: *{cfg('usd_rate')} BDT*\n"
        f"👤 Support: *{cfg('support_username')}*\n"
        f"🎁 Daily Bonus: *{cfg('daily_bonus_amount')} BDT*\n"
        f"🤖 Supplier Bot: *{supplier}*\n"
        f"📢 Channels: *{len(FORCE_JOIN_CHANNELS())}*\n"
        f"💳 Payment Methods: *{len(PAYMENT_METHODS())}*\n"
        "✨━━━━━━━━━━━━✨"
    )
    try:
        bot.edit_message_text(txt, ADMIN_ID, call.message.message_id,
                              reply_markup=mk, parse_mode="Markdown")
    except Exception:
        bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("cfg|"))
def config_router(call):
    if call.message.chat.id != ADMIN_ID: return
    key = call.data.split("|")[1]; bot.answer_callback_query(call.id)
    if key == "usd_rate":
        msg = bot.send_message(ADMIN_ID, f"💵 Current: *{cfg('usd_rate')} BDT*\nSend new rate:", parse_mode="Markdown")
        bot.register_next_step_handler(msg, _cfg_save_float, "usd_rate")
    elif key == "support_username":
        msg = bot.send_message(ADMIN_ID, f"👤 Current: *{cfg('support_username')}*\nSend new username:", parse_mode="Markdown")
        bot.register_next_step_handler(msg, _cfg_save_str, "support_username")
    elif key == "daily_bonus_amount":
        msg = bot.send_message(ADMIN_ID, f"🎁 Current: *{cfg('daily_bonus_amount')} BDT*\nSend new amount:", parse_mode="Markdown")
        bot.register_next_step_handler(msg, _cfg_save_float, "daily_bonus_amount")
    elif key == "maintenance_message":
        msg = bot.send_message(ADMIN_ID, "🔧 Send new maintenance message:", parse_mode="Markdown")
        bot.register_next_step_handler(msg, _cfg_save_str, "maintenance_message")
    elif key == "supplier_bot":
        cur = cfg("supplier_bot_username") or "(not set)"
        msg = bot.send_message(ADMIN_ID,
            f"🤖 *Supplier Bot Username*\nCurrent: `{cur}`\n\nSend new username (e.g. `@supplier_bot`):",
            parse_mode="Markdown")
        bot.register_next_step_handler(msg, _cfg_save_supplier)
    elif key == "channels":    _adm_channels(call)
    elif key == "payments":    _adm_payments(call)
    elif key == "back_config": _adm_config(call)

def _cfg_save_float(message, key):
    if message.chat.id != ADMIN_ID: return
    try:
        val = float(message.text.strip()); settings[key] = val; save_settings(settings)
        bot.send_message(ADMIN_ID, f"✅ *{key}* → `{val}`", parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, "❌ Invalid value. Enter a number.")

def _cfg_save_str(message, key):
    if message.chat.id != ADMIN_ID: return
    val = message.text.strip(); settings[key] = val; save_settings(settings)
    bot.send_message(ADMIN_ID, f"✅ *{key}* updated!", parse_mode="Markdown")

def _cfg_save_supplier(message):
    if message.chat.id != ADMIN_ID: return
    val = message.text.strip()
    if not val.startswith("@"):
        val = "@" + val
    settings["supplier_bot_username"] = val
    save_settings(settings)
    bot.send_message(ADMIN_ID,
        f"✅ *Supplier Bot* set to `{val}`\n\nSync Stock and VPN delivery are now active.",
        parse_mode="Markdown")


# ─────────────────────────────────────────────
#  Channel Links Manager
# ─────────────────────────────────────────────
def _adm_channels(call):
    channels = FORCE_JOIN_CHANNELS()
    mk = types.InlineKeyboardMarkup(row_width=2)
    btns = [types.InlineKeyboardButton(f"🗑️ {ch['name']}", callback_data=f"ch_del|{i}")
            for i, ch in enumerate(channels)]
    if btns: mk.add(*btns)
    mk.add(
        types.InlineKeyboardButton("➕ Add Channel", callback_data="ch_add"),
        types.InlineKeyboardButton("🔙 Back",        callback_data="cfg|back_config"),
    )
    lines = "\n".join(f"  • {ch['name']} (`{ch['username']}`)" for ch in channels) or "  _(none)_"
    txt   = f"🔗 *Force Join Channels*\n✨━━━━━━━━━━━━✨\n{lines}\n✨━━━━━━━━━━━━✨"
    try: bot.edit_message_text(txt, ADMIN_ID, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")

@bot.callback_query_handler(func=lambda c: c.data == "ch_add")
def ch_add_start(call):
    if call.message.chat.id != ADMIN_ID: return
    bot.answer_callback_query(call.id)
    msg = bot.send_message(ADMIN_ID,
        "➕ *Add Force Join Channel*\nFormat: `@username|Display Name`\nExample: `@prime_bazar|Prime Bazar`",
        parse_mode="Markdown")
    bot.register_next_step_handler(msg, ch_add_save)

def ch_add_save(message):
    if message.chat.id != ADMIN_ID: return
    try:
        parts = message.text.strip().split("|"); username = parts[0].strip(); name = parts[1].strip()
        url   = f"https://t.me/{username.lstrip('@')}"
        channels = settings.get("force_join_channels", [])
        channels.append({"username": username, "url": url, "name": name})
        settings["force_join_channels"] = channels; save_settings(settings)
        bot.send_message(ADMIN_ID, f"✅ Channel *{name}* added!", parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, "❌ Wrong format. Use: `@username|Display Name`", parse_mode="Markdown")

@bot.callback_query_handler(func=lambda c: c.data.startswith("ch_del|"))
def ch_del(call):
    if call.message.chat.id != ADMIN_ID: return
    idx = int(call.data.split("|")[1]); channels = settings.get("force_join_channels", [])
    if 0 <= idx < len(channels):
        removed = channels.pop(idx); settings["force_join_channels"] = channels; save_settings(settings)
        bot.answer_callback_query(call.id, f"✅ Removed: {removed['name']}")
        _adm_channels(call)
    else: bot.answer_callback_query(call.id, "❌ Not found.")


# ─────────────────────────────────────────────
#  Payment Methods Manager
# ─────────────────────────────────────────────
def _adm_payments(call):
    methods = PAYMENT_METHODS()
    min_bdt = cfg("min_deposit_bdt") or 20
    min_usd = cfg("min_deposit_usd") or 1
    mk      = types.InlineKeyboardMarkup(row_width=2)
    for method in methods:
        mk.add(
            types.InlineKeyboardButton(f"✏️ {method}", callback_data=f"pay_edit|{method}"),
            types.InlineKeyboardButton(f"🗑️ Delete",   callback_data=f"pay_del|{method}"),
        )
    mk.add(
        types.InlineKeyboardButton("➕ Add Method",         callback_data="pay_add"),
        types.InlineKeyboardButton("🔙 Back",               callback_data="cfg|back_config"),
    )
    mk.add(
        types.InlineKeyboardButton(f"💵 Min BDT: {min_bdt} BDT", callback_data="pay_min|bdt"),
        types.InlineKeyboardButton(f"💲 Min USD: ${min_usd}",    callback_data="pay_min|usd"),
    )
    txt = (
        f"💳 *Payment Methods*\n✨━━━━━━━━━━━━✨\n"
        f"Tap to edit or delete a method.\n\n"
        f"⚙️ *Minimum Deposit:*\n"
        f"  🟢 BDT: *{min_bdt} BDT*\n"
        f"  🟢 USD (Binance): *${min_usd}*\n"
        f"✨━━━━━━━━━━━━✨"
    )
    try: bot.edit_message_text(txt, ADMIN_ID, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")

@bot.callback_query_handler(func=lambda c: c.data.startswith("pay_min|"))
def pay_min_start(call):
    if call.message.chat.id != ADMIN_ID: return
    bot.answer_callback_query(call.id)
    kind = call.data.split("|")[1]   # "bdt" or "usd"
    if kind == "bdt":
        cur = cfg("min_deposit_bdt") or 20
        msg = bot.send_message(ADMIN_ID,
            f"💵 *Minimum BDT Deposit*\nবর্তমান: *{cur} BDT*\n\nনতুন পরিমাণ লিখুন (শুধু সংখ্যা):",
            parse_mode="Markdown")
        bot.register_next_step_handler(msg, _pay_min_save, "min_deposit_bdt")
    else:
        cur = cfg("min_deposit_usd") or 1
        msg = bot.send_message(ADMIN_ID,
            f"💲 *Minimum USD Deposit (Binance)*\nবর্তমান: *${cur}*\n\nনতুন পরিমাণ লিখুন (শুধু সংখ্যা):",
            parse_mode="Markdown")
        bot.register_next_step_handler(msg, _pay_min_save, "min_deposit_usd")

def _pay_min_save(message, key):
    if message.chat.id != ADMIN_ID: return
    try:
        val = float(message.text.strip())
        if val <= 0: raise ValueError
        settings[key] = val; save_settings(settings)
        label = "BDT" if key == "min_deposit_bdt" else "USD"
        sym   = "" if key == "min_deposit_bdt" else "$"
        bot.send_message(ADMIN_ID,
            f"✅ Minimum {label} Deposit আপডেট হয়েছে: *{sym}{val} {label}*",
            parse_mode="Markdown")
    except Exception:
        bot.send_message(ADMIN_ID, "❌ ভুল ইনপুট। শুধু সংখ্যা দিন (যেমন: 50)")


@bot.callback_query_handler(func=lambda c: c.data.startswith("pay_edit|"))
def pay_edit_start(call):
    if call.message.chat.id != ADMIN_ID: return
    method = call.data.split("|")[1]; bot.answer_callback_query(call.id)
    msg = bot.send_message(ADMIN_ID,
        f"✏️ New value for *{method}*\nCurrent: `{PAYMENT_METHODS().get(method,'')}`\n\nSend new value:",
        parse_mode="Markdown")
    bot.register_next_step_handler(msg, pay_edit_save, method)

def pay_edit_save(message, method):
    if message.chat.id != ADMIN_ID: return
    val = message.text.strip(); settings["payment_methods"][method] = val; save_settings(settings)
    bot.send_message(ADMIN_ID, f"✅ *{method}* updated to:\n`{val}`", parse_mode="Markdown")

@bot.callback_query_handler(func=lambda c: c.data.startswith("pay_del|"))
def pay_del(call):
    if call.message.chat.id != ADMIN_ID: return
    method = call.data.split("|")[1]
    if method in settings.get("payment_methods", {}):
        del settings["payment_methods"][method]; save_settings(settings)
        bot.answer_callback_query(call.id, f"✅ {method} deleted!")
        _adm_payments(call)
    else: bot.answer_callback_query(call.id, "❌ Not found.")

@bot.callback_query_handler(func=lambda c: c.data == "pay_add")
def pay_add_start(call):
    if call.message.chat.id != ADMIN_ID: return
    bot.answer_callback_query(call.id)
    msg = bot.send_message(ADMIN_ID,
        "➕ *Add Payment Method*\nFormat: `MethodName|Number`\nExample: `Nagad|01712345678`",
        parse_mode="Markdown")
    bot.register_next_step_handler(msg, pay_add_save)

def pay_add_save(message):
    if message.chat.id != ADMIN_ID: return
    try:
        parts = message.text.strip().split("|"); name = parts[0].strip(); detail = parts[1].strip()
        settings.setdefault("payment_methods", {})[name] = detail; save_settings(settings)
        bot.send_message(ADMIN_ID, f"✅ *{name}* added!", parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, "❌ Wrong format. Use: `MethodName|Number`", parse_mode="Markdown")


# ─────────────────────────────────────────────
#  Feature Toggles
# ─────────────────────────────────────────────
def _adm_toggles(call):
    fj = cfg("force_join_enabled"); ls = cfg("language_selection_enabled")
    db = cfg("daily_bonus_enabled"); mm = cfg("maintenance_mode")
    ub = cfg("userbot_enabled") is not False
    mk = types.InlineKeyboardMarkup(row_width=2)
    mk.add(
        types.InlineKeyboardButton(f"{'✅' if fj else '❌'} Force Join",  callback_data="toggle|force_join_enabled"),
        types.InlineKeyboardButton(f"{'✅' if ls else '❌'} Lang Select", callback_data="toggle|language_selection_enabled"),
    )
    mk.add(
        types.InlineKeyboardButton(f"{'✅' if db else '❌'} Daily Bonus", callback_data="toggle|daily_bonus_enabled"),
        types.InlineKeyboardButton(f"{'🔧' if mm else '✅'} Maintenance", callback_data="toggle|maintenance_mode"),
    )
    mk.add(
        types.InlineKeyboardButton(f"{'✅' if ub else '❌'} Userbot", callback_data="toggle|userbot_enabled"),
    )
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="adm|back"))
    txt = (
        f"🔀 *Feature Toggles*\n✨━━━━━━━━━━━━✨\n"
        f"📢 Force Join: {'✅ ON' if fj else '❌ OFF'}\n"
        f"🌐 Language Select: {'✅ ON' if ls else '❌ OFF'}\n"
        f"🎁 Daily Bonus: {'✅ ON' if db else '❌ OFF'}\n"
        f"🔧 Maintenance: {'🔧 ON' if mm else '✅ OFF'}\n"
        f"🤖 Userbot: {'✅ ON' if ub else '❌ OFF'}\n"
        f"✨━━━━━━━━━━━━✨\n"
        f"_⚠️ Userbot OFF করলে সব VPN Manual Delivery হবে।_\n"
        f"_Tap any button to toggle._"
    )
    try: bot.edit_message_text(txt, ADMIN_ID, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")

@bot.callback_query_handler(func=lambda c: c.data.startswith("toggle|"))
def feature_toggle(call):
    if call.message.chat.id != ADMIN_ID: return
    key = call.data.split("|")[1]; settings[key] = not settings.get(key, False); save_settings(settings)
    state = "✅ ON" if settings[key] else "❌ OFF"
    bot.answer_callback_query(call.id, f"{key} → {state}")
    _adm_toggles(call)


# ─────────────────────────────────────────────
#  Coupon Manager
# ─────────────────────────────────────────────
def _adm_manual_orders(call):
    """Admin panel — list all pending manual VPN deliveries with Send/Reject actions."""
    orders = _pending_manual_orders
    mk = types.InlineKeyboardMarkup(row_width=2)
    if orders:
        for uid, o in list(orders.items()):
            u_name = o.get("user_name", "?")
            d_name = o.get("d_name", "?")
            mk.add(
                types.InlineKeyboardButton(f"📨 Send · {u_name}", callback_data=f"adm_send_vpn|{uid}"),
                types.InlineKeyboardButton(f"❌ Reject · {u_name}", callback_data=f"adm_reject_vpn|{uid}"),
            )
    else:
        mk.add(types.InlineKeyboardButton("✅ কোনো pending order নেই", callback_data="noop"))
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="adm|back"))
    txt = (
        f"📋 *Pending Manual VPN Orders*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"🔔 Total Pending: *{len(orders)}*\n\n"
    )
    for uid, o in list(orders.items()):
        txt += (
            f"👤 `{uid}` — {o.get('user_name','?')}\n"
            f"   🛡️ {o.get('d_name','?')} | ⏳ {o.get('duration','?')}\n"
            f"   💰 {o.get('total','?')} BDT | 🕐 {o.get('time','?')}\n\n"
        )
    txt += "✨━━━━━━━━━━━━━━━━━━✨\n_Send করে VPN দিন, বা ভুল/স্প্যাম অর্ডার হলে Reject করুন (টাকা রিফান্ড হয়ে যাবে)।_"
    try:
        bot.edit_message_text(txt, ADMIN_ID, call.message.message_id,
                              reply_markup=mk, parse_mode="Markdown")
    except Exception:
        bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("adm_reject_vpn|"))
def adm_reject_vpn_confirm(call):
    """Admin clicks 'Reject' on a manual order — ask for confirmation first."""
    if call.message.chat.id != ADMIN_ID: return
    target_uid = call.data.split("|")[1]
    order = _pending_manual_orders.get(target_uid)
    bot.answer_callback_query(call.id)
    if not order:
        bot.send_message(ADMIN_ID, "❌ এই অর্ডারটি আর নেই (আগেই প্রসেস হয়ে গেছে)।"); return
    mk = types.InlineKeyboardMarkup(row_width=1)
    mk.add(
        types.InlineKeyboardButton("✅ হ্যাঁ, Reject করে রিফান্ড দিন", callback_data=f"adm_reject_ok|{target_uid}"),
        types.InlineKeyboardButton("🔙 না, ফিরে যান", callback_data="adm|manual_orders"),
    )
    bot.send_message(
        ADMIN_ID,
        f"⚠️ *Reject Confirmation*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"👤 User: `{target_uid}` — {order.get('user_name','?')}\n"
        f"🛡️ Product: *{order.get('d_name','?')}* | ⏳ {order.get('duration','?')}\n"
        f"💰 Amount: *{order.get('total','?')} BDT*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"এই অর্ডারটি Reject করলে ইউজারকে *{order.get('total','?')} BDT* রিফান্ড করে দেওয়া হবে। নিশ্চিত?",
        reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("adm_reject_ok|"))
def adm_reject_vpn_do(call):
    """Admin confirmed reject — refund balance, notify user, drop the pending order."""
    if call.message.chat.id != ADMIN_ID: return
    target_uid = call.data.split("|")[1]
    bot.answer_callback_query(call.id)
    order = _pending_manual_orders.pop(target_uid, None)
    _save_manual_orders(_pending_manual_orders)
    if not order:
        bot.send_message(ADMIN_ID, "❌ এই অর্ডারটি আর নেই (আগেই প্রসেস হয়ে গেছে)।"); return
    total = order.get("total", 0)
    try:
        update_user(target_uid, "balance", total)
    except Exception:
        pass
    new_bal = round(get_user(target_uid).get("balance", 0), 2)
    try:
        bot.send_message(
            target_uid,
            f"❌ *আপনার অর্ডারটি বাতিল করা হয়েছে!*\n"
            f"✨━━━━━━━━━━━━━━━━━━✨\n"
            f"🛡️ Product: *{order.get('d_name','?')}* | ⏳ {order.get('duration','?')}\n"
            f"💰 Refunded: *{total} BDT*\n"
            f"💳 Current Balance: *{new_bal} BDT*\n"
            f"✨━━━━━━━━━━━━━━━━━━✨\n"
            f"কোনো সমস্যা মনে হলে Support-এ যোগাযোগ করুন।",
            parse_mode="Markdown")
    except Exception:
        pass
    bot.send_message(
        ADMIN_ID,
        f"✅ *Order Rejected & Refunded!*\n"
        f"👤 User `{target_uid}` কে *{total} BDT* রিফান্ড করা হয়েছে।\n"
        f"💳 New Balance: *{new_bal} BDT*",
        parse_mode="Markdown")


def _adm_coupons(call):
    global coupons; coupons = load_coupons()
    mk = types.InlineKeyboardMarkup(row_width=2)
    btns = [types.InlineKeyboardButton(
                f"🎟️ {code} ({c.get('uses_left',0)} left)", callback_data=f"cpn_view|{code}")
            for code, c in coupons.items()]
    if btns: mk.add(*btns)
    mk.add(
        types.InlineKeyboardButton("➕ Create Coupon", callback_data="cpn_add"),
        types.InlineKeyboardButton("🔙 Back",          callback_data="adm|back"),
    )
    txt = f"🎟️ *Coupon Manager*\n✨━━━━━━━━━━━━✨\nTotal: *{len(coupons)}*"
    try: bot.edit_message_text(txt, ADMIN_ID, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")

@bot.callback_query_handler(func=lambda c: c.data == "cpn_add")
def cpn_add_start(call):
    if call.message.chat.id != ADMIN_ID: return
    bot.answer_callback_query(call.id)
    msg = bot.send_message(ADMIN_ID,
        "➕ *Create Coupon*\nFormat: `CODE|amount_BDT|max_uses`\nExample: `WELCOME50|50|100`",
        parse_mode="Markdown")
    bot.register_next_step_handler(msg, cpn_add_save)

def cpn_add_save(message):
    if message.chat.id != ADMIN_ID: return
    try:
        parts = message.text.strip().split("|"); code = parts[0].strip().upper()
        amount = float(parts[1].strip()); uses = int(parts[2].strip())
        global coupons; coupons = load_coupons()
        coupons[code] = {"amount": amount, "uses_left": uses, "used_by": []}; save_coupons(coupons)
        bot.send_message(ADMIN_ID,
            f"✅ Coupon *{code}* created!\n💰 *{amount} BDT* | 🔢 Max uses: *{uses}*",
            parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, "❌ Wrong format. Use: `CODE|amount|max_uses`", parse_mode="Markdown")

@bot.callback_query_handler(func=lambda c: c.data.startswith("cpn_view|"))
def cpn_view(call):
    if call.message.chat.id != ADMIN_ID: return
    code = call.data.split("|")[1]; global coupons; coupons = load_coupons()
    c = coupons.get(code)
    if not c: bot.answer_callback_query(call.id, "❌ Not found."); return
    mk = types.InlineKeyboardMarkup(row_width=2)
    mk.add(
        types.InlineKeyboardButton("🗑️ Delete", callback_data=f"cpn_del|{code}"),
        types.InlineKeyboardButton("🔙 Back",   callback_data="adm|coupons"),
    )
    txt = (
        f"🎟️ *Coupon: {code}*\n✨━━━━━━━━━━━━✨\n"
        f"💰 Amount: *{c.get('amount',0)} BDT*\n"
        f"🔢 Uses Left: *{c.get('uses_left',0)}*\n"
        f"👥 Used by: *{len(c.get('used_by',[]))} users*\n✨━━━━━━━━━━━━✨"
    )
    try: bot.edit_message_text(txt, ADMIN_ID, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")
    bot.answer_callback_query(call.id)

@bot.callback_query_handler(func=lambda c: c.data.startswith("cpn_del|"))
def cpn_delete(call):
    if call.message.chat.id != ADMIN_ID: return
    code = call.data.split("|")[1]; global coupons; coupons = load_coupons()
    if code in coupons:
        del coupons[code]; save_coupons(coupons)
        bot.answer_callback_query(call.id, f"✅ {code} deleted!")
        _adm_coupons(call)
    else: bot.answer_callback_query(call.id, "❌ Not found.")


# ─────────────────────────────────────────────
#  VPN Duration Management
# ─────────────────────────────────────────────
def _adm_vpndurs(call):
    durs = get_vpn_durations()
    mk   = types.InlineKeyboardMarkup(row_width=2)
    btns = [types.InlineKeyboardButton(f"✏️ {di['name']}", callback_data=f"vpndur_edit|{dk}")
            for dk, di in durs.items()]
    if btns: mk.add(*btns)
    mk.add(
        types.InlineKeyboardButton("➕ Add Duration", callback_data="vpndur_add"),
        types.InlineKeyboardButton("🔙 Back",         callback_data="adm|back"),
    )
    try:
        bot.edit_message_text(
            "🛡️ *VPN Duration Management*\n💎━━━━━━━━━━━━━━━━━━💎\nTap to rename or delete.",
            ADMIN_ID, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    except Exception:
        bot.send_message(ADMIN_ID,
            "🛡️ *VPN Duration Management*\n💎━━━━━━━━━━━━━━━━━━💎",
            reply_markup=mk, parse_mode="Markdown")
    bot.answer_callback_query(call.id)

@bot.callback_query_handler(func=lambda c: c.data == "vpndur_add")
def vpndur_add_start(call):
    if call.message.chat.id != ADMIN_ID: return
    bot.answer_callback_query(call.id)
    msg = bot.send_message(ADMIN_ID,
        "🛡️ *Add VPN Duration*\nFormat: `key|emoji Name`\nExample: `03d|⚡ 03 Days`",
        parse_mode="Markdown")
    bot.register_next_step_handler(msg, vpndur_add_save)

def vpndur_add_save(message):
    if message.chat.id != ADMIN_ID: return
    try:
        key, name = message.text.strip().split("|", 1)
        key = key.strip().lower().replace(" ", "_"); name = name.strip()
        market_data["vpn_durations"][key] = {"name": name}; save_market_data(market_data)
        bot.send_message(ADMIN_ID, f"✅ Duration *{name}* added!", parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, "❌ Wrong format. Use: `key|emoji Name`", parse_mode="Markdown")

@bot.callback_query_handler(func=lambda c: c.data.startswith("vpndur_edit|"))
def vpndur_edit_menu(call):
    if call.message.chat.id != ADMIN_ID: return
    dk = call.data.split("|")[1]; durs = get_vpn_durations(); dur_name = durs.get(dk, {}).get("name", dk)
    mk = types.InlineKeyboardMarkup(row_width=2)
    mk.add(
        types.InlineKeyboardButton("✏️ Rename", callback_data=f"vpndur_rename|{dk}"),
        types.InlineKeyboardButton("🗑️ Delete", callback_data=f"vpndur_del|{dk}"),
    )
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="adm|vpndurs"))
    bot.edit_message_text(f"🛡️ *{dur_name}*\n💎━━━━━━━━━━━━💎\nKey: `{dk}`\nChoose action:",
                          ADMIN_ID, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    bot.answer_callback_query(call.id)

@bot.callback_query_handler(func=lambda c: c.data.startswith("vpndur_rename|"))
def vpndur_rename_start(call):
    if call.message.chat.id != ADMIN_ID: return
    dk = call.data.split("|")[1]; bot.answer_callback_query(call.id)
    msg = bot.send_message(ADMIN_ID, f"✏️ New name for `{dk}` (e.g. `⚡ 03 Days`):", parse_mode="Markdown")
    bot.register_next_step_handler(msg, vpndur_rename_save, dk)

def vpndur_rename_save(message, dk):
    if message.chat.id != ADMIN_ID: return
    new_name = message.text.strip()
    if dk in market_data.get("vpn_durations", {}):
        market_data["vpn_durations"][dk]["name"] = new_name; save_market_data(market_data)
        bot.send_message(ADMIN_ID, f"✅ Renamed to *{new_name}*!", parse_mode="Markdown")
    else: bot.send_message(ADMIN_ID, "❌ Duration not found.")

@bot.callback_query_handler(func=lambda c: c.data.startswith("vpndur_del|"))
def vpndur_delete(call):
    if call.message.chat.id != ADMIN_ID: return
    dk = call.data.split("|")[1]
    if dk in market_data.get("vpn_durations", {}):
        del market_data["vpn_durations"][dk]; save_market_data(market_data)
        bot.answer_callback_query(call.id, "✅ Deleted!")
        _adm_vpndurs(call)
    else: bot.answer_callback_query(call.id, "❌ Not found.")


# ─────────────────────────────────────────────
#  Category Management
# ─────────────────────────────────────────────
def _adm_cats(call):
    cats = get_categories()
    mk   = types.InlineKeyboardMarkup(row_width=2)
    for cat_key, cat_info in cats.items():
        mk.add(
            types.InlineKeyboardButton(f"{cat_info['emoji']} {cat_info['name']}", callback_data=f"cat_rename|{cat_key}"),
            types.InlineKeyboardButton("🗑️ Delete", callback_data=f"cat_del|{cat_key}"),
        )
    mk.add(
        types.InlineKeyboardButton("➕ Add Category", callback_data="cat_add"),
        types.InlineKeyboardButton("🔙 Back",         callback_data="adm|back"),
    )
    try:
        bot.edit_message_text(
            "📦 *Category Management*\n✨━━━━━━━━━━━━✨\nTap name to rename, 🗑️ to delete.",
            ADMIN_ID, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    except Exception:
        bot.send_message(ADMIN_ID,
            "📦 *Category Management*\n✨━━━━━━━━━━━━✨",
            reply_markup=mk, parse_mode="Markdown")
    bot.answer_callback_query(call.id)

@bot.callback_query_handler(func=lambda c: c.data == "cat_add")
def cat_add_start(call):
    if call.message.chat.id != ADMIN_ID: return
    bot.answer_callback_query(call.id)
    msg = bot.send_message(ADMIN_ID,
        "📦 *New Category*\nFormat: `key|emoji|Full Name`\nExample: `food|🍕|Food & Drinks`",
        parse_mode="Markdown")
    bot.register_next_step_handler(msg, cat_add_save)

def cat_add_save(message):
    if message.chat.id != ADMIN_ID: return
    try:
        parts = message.text.strip().split("|")
        key = parts[0].strip().lower().replace(" ","_"); emoji = parts[1].strip(); name = parts[2].strip()
        market_data["categories"][key] = {"name": name, "emoji": emoji}; save_market_data(market_data)
        bot.send_message(ADMIN_ID, f"✅ Category *{name}* added!", parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, "❌ Wrong format. Use: `key|emoji|Full Name`", parse_mode="Markdown")

@bot.callback_query_handler(func=lambda c: c.data.startswith("cat_rename|"))
def cat_rename_start(call):
    if call.message.chat.id != ADMIN_ID: return
    cat_key = call.data.split("|")[1]; cats = get_categories()
    if cat_key not in cats: bot.answer_callback_query(call.id, "❌ Not found."); return
    bot.answer_callback_query(call.id)
    ci = cats[cat_key]
    msg = bot.send_message(ADMIN_ID,
        f"✏️ *Rename Category*\nCurrent: {ci['emoji']} *{ci['name']}*\n\nFormat: `emoji|New Name`",
        parse_mode="Markdown")
    bot.register_next_step_handler(msg, cat_rename_save, cat_key)

def cat_rename_save(message, cat_key):
    if message.chat.id != ADMIN_ID: return
    try:
        parts = message.text.strip().split("|")
        market_data["categories"][cat_key]["emoji"] = parts[0].strip()
        market_data["categories"][cat_key]["name"]  = parts[1].strip()
        save_market_data(market_data)
        bot.send_message(ADMIN_ID, "✅ Category renamed!", parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, "❌ Wrong format. Use: `emoji|New Name`", parse_mode="Markdown")

@bot.callback_query_handler(func=lambda c: c.data.startswith("cat_del|"))
def cat_delete(call):
    if call.message.chat.id != ADMIN_ID: return
    cat_key = call.data.split("|")[1]
    if cat_key in get_categories():
        del market_data["categories"][cat_key]; save_market_data(market_data)
        bot.answer_callback_query(call.id, "✅ Deleted!")
        _adm_cats(call)
    else: bot.answer_callback_query(call.id, "❌ Not found.")


# ─────────────────────────────────────────────
#  Product Management
# ─────────────────────────────────────────────
def _adm_prods_cat_select(call):
    cats = get_categories()
    mk   = types.InlineKeyboardMarkup(row_width=2)
    btns = [types.InlineKeyboardButton(cat_info["name"], callback_data=f"prod_cat|{cat_key}")
            for cat_key, cat_info in cats.items()]
    if btns: mk.add(*btns)
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="adm|back"))
    try:
        bot.edit_message_text("🛍️ *Product Management*\n✨━━━━━━━━━━━━✨\nSelect a category:",
                              ADMIN_ID, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    except Exception:
        bot.send_message(ADMIN_ID, "🛍️ *Product Management*\n✨━━━━━━━━━━━━✨\nSelect a category:",
                         reply_markup=mk, parse_mode="Markdown")
    bot.answer_callback_query(call.id)

@bot.callback_query_handler(func=lambda c: c.data.startswith("prod_cat|"))
def prod_list(call):
    if call.message.chat.id != ADMIN_ID: return
    cat_key  = call.data.split("|")[1]; prods = get_products(); cats = get_categories()
    cat_name = cats.get(cat_key, {}).get("name", cat_key)
    mk   = types.InlineKeyboardMarkup(row_width=2)
    btns = [types.InlineKeyboardButton(
                f"✏️ {p_name} — {p_info['price']} BDT", callback_data=f"prod_edit|{p_name}")
            for p_name, p_info in prods.items() if p_info.get("cat") == cat_key]
    if btns: mk.add(*btns)
    mk.add(
        types.InlineKeyboardButton(f"➕ Add Product", callback_data=f"prod_add|{cat_key}"),
        types.InlineKeyboardButton("🔙 Back",          callback_data="adm|prods"),
    )
    bot.edit_message_text(f"🛍️ *{cat_name}* — Products\n✨━━━━━━━━━━━━✨",
                          ADMIN_ID, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    bot.answer_callback_query(call.id)

@bot.callback_query_handler(func=lambda c: c.data.startswith("prod_add|"))
def prod_add_start(call):
    if call.message.chat.id != ADMIN_ID: return
    cat_key = call.data.split("|")[1]; bot.answer_callback_query(call.id)
    if cat_key == "vpn":
        durs = get_vpn_durations()
        dur_list = "\n".join(f"  `{k}` → {v['name']}" for k, v in durs.items())
        msg = bot.send_message(ADMIN_ID,
            f"🛡️ *Add VPN Product*\nFormat: `Name|Duration Label|Price|dur_key`\n"
            f"Available keys:\n{dur_list}\nExample: `Nord VPN|30 Days|300|30d`",
            parse_mode="Markdown")
    else:
        msg = bot.send_message(ADMIN_ID,
            "🛍️ *Add Product*\nFormat: `Name|Duration|Price`\nExample: `Gmail|Lifetime|1.5`",
            parse_mode="Markdown")
    bot.register_next_step_handler(msg, prod_add_save, cat_key)

def prod_add_save(message, cat_key):
    if message.chat.id != ADMIN_ID: return
    try:
        parts = message.text.strip().split("|")
        name = parts[0].strip(); duration = parts[1].strip(); price = float(parts[2].strip())
        filename = name.lower().replace(" ", "_") + ".xlsx"
        key = _unique_product_key(name)
        entry = {"name": name, "file": filename, "price": price, "duration": duration, "cat": cat_key}
        if cat_key == "vpn":
            if len(parts) >= 4:
                entry["vpn_dur"] = parts[3].strip()
            else:
                entry["vpn_dur"] = guess_vpn_dur_key(duration)
        market_data["products"][key] = entry; save_market_data(market_data)
        key_info = f"\n🔑 Key: `{key}`" if key != name else ""
        bot.send_message(ADMIN_ID,
            f"✅ *{name}* added!{key_info}\n💸 {price} BDT | ⏱️ {duration}\n📁 `{filename}`",
            parse_mode="Markdown")
    except Exception:
        fmt = "`Name|Duration|Price|dur_key`" if cat_key == "vpn" else "`Name|Duration|Price`"
        bot.send_message(ADMIN_ID, f"❌ Wrong format. Use: {fmt}", parse_mode="Markdown")

@bot.callback_query_handler(func=lambda c: c.data.startswith("prod_edit|"))
def prod_edit_menu(call):
    if call.message.chat.id != ADMIN_ID: return
    p_name = call.data.split("|", 1)[1]; prods = get_products()
    if p_name not in prods: bot.answer_callback_query(call.id, "❌ Product not found."); return
    p = prods[p_name]; stock = get_stock_count(p_name)
    is_vpn = p.get("cat") == "vpn"
    mk = types.InlineKeyboardMarkup(row_width=2)
    mk.add(
        types.InlineKeyboardButton("📝 Edit Name",      callback_data=f"prod_ename|{p_name}"),
        types.InlineKeyboardButton("💰 Edit Price",     callback_data=f"prod_eprice|{p_name}"),
    )
    mk.add(
        types.InlineKeyboardButton("⏳ Edit Duration",  callback_data=f"prod_edur|{p_name}"),
        types.InlineKeyboardButton("🗑️ Delete",         callback_data=f"prod_del|{p_name}"),
    )
    # VPN-only: per-product Manual/Auto delivery toggle
    if is_vpn:
        per_toggle = p.get("manual_delivery")   # None = global rule, True = manual, False = auto
        if per_toggle is True:
            tog_label = "🔴 Delivery: Manual (tap to → Auto)"
        elif per_toggle is False:
            tog_label = "🟢 Delivery: Auto (tap to → Manual)"
        else:
            # inherits global setting — show current effective value
            effective = _is_manual_delivery_vpn(p, p_name)
            tog_label = f"{'🔴 Manual' if effective else '🟢 Auto'} (Global — tap to set)"
        mk.add(types.InlineKeyboardButton(tog_label, callback_data=f"prod_tog_manual|{p_name}"))
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data=f"prod_cat|{p['cat']}"))
    # Build info text
    if is_vpn:
        per_toggle = p.get("manual_delivery")
        if per_toggle is True:   delivery_txt = "🔴 Manual (per-product)"
        elif per_toggle is False: delivery_txt = "🟢 Auto/Userbot (per-product)"
        else:                    delivery_txt = "🌐 Inherits global rule"
    else:
        delivery_txt = "N/A"
    txt = (
        f"🛍️ *{p_name}*\n✨━━━━━━━━━━━━✨\n"
        f"💰 Price: *{p['price']} BDT*\n"
        f"⏳ Duration: *{p.get('duration','N/A')}*\n"
        f"📦 Stock: *{stock} pcs*\n"
        + (f"🚚 Delivery: {delivery_txt}\n" if is_vpn else "")
        + f"✨━━━━━━━━━━━━✨"
    )
    bot.edit_message_text(txt, ADMIN_ID, call.message.message_id, reply_markup=mk, parse_mode="Markdown")
    bot.answer_callback_query(call.id)


@bot.callback_query_handler(func=lambda c: c.data.startswith("prod_tog_manual|"))
def prod_toggle_manual_delivery(call):
    """Toggle per-product manual delivery: None → True → False → None (cycle)."""
    if call.message.chat.id != ADMIN_ID: return
    p_name = call.data.split("|", 1)[1]
    prods  = get_products()
    if p_name not in prods:
        bot.answer_callback_query(call.id, "❌ Product not found."); return
    current = prods[p_name].get("manual_delivery")  # None / True / False
    # Cycle: None → True (Manual) → False (Auto) → None (Global)
    if current is None:
        new_val = True;  label = "🔴 Manual Delivery ON"
    elif current is True:
        new_val = False; label = "🟢 Auto Delivery ON"
    else:
        new_val = None;  label = "🌐 Reverted to Global Rule"
    if new_val is None:
        market_data["products"][p_name].pop("manual_delivery", None)
    else:
        market_data["products"][p_name]["manual_delivery"] = new_val
    save_market_data(market_data)
    bot.answer_callback_query(call.id, f"✅ {label}")
    # Refresh the product edit menu
    prod_edit_menu(call)

@bot.callback_query_handler(func=lambda c: c.data.startswith("prod_eprice|"))
def prod_edit_price_start(call):
    if call.message.chat.id != ADMIN_ID: return
    p_name = call.data.split("|", 1)[1]; bot.answer_callback_query(call.id)
    msg = bot.send_message(ADMIN_ID, f"💸 New price for *{p_name}* (BDT):", parse_mode="Markdown")
    bot.register_next_step_handler(msg, prod_edit_price_save, p_name)

def prod_edit_price_save(message, p_name):
    if message.chat.id != ADMIN_ID: return
    try:
        new_price = float(message.text.strip())
        market_data["products"][p_name]["price"] = new_price; save_market_data(market_data)
        bot.send_message(ADMIN_ID, f"✅ *{p_name}* price → *{new_price} BDT*!", parse_mode="Markdown")
    except Exception: bot.send_message(ADMIN_ID, "❌ Invalid price. Enter a number.", parse_mode="Markdown")

@bot.callback_query_handler(func=lambda c: c.data.startswith("prod_edur|"))
def prod_edit_dur_start(call):
    if call.message.chat.id != ADMIN_ID: return
    p_name = call.data.split("|", 1)[1]; p = get_products().get(p_name, {}); bot.answer_callback_query(call.id)
    if p.get("cat") == "vpn":
        durs = get_vpn_durations()
        mk   = types.InlineKeyboardMarkup(row_width=2)
        btns = [types.InlineKeyboardButton(di["name"], callback_data=f"prod_setdurkey|{p_name}|{dk}")
                for dk, di in durs.items()]
        if btns: mk.add(*btns)
        mk.add(types.InlineKeyboardButton("🔙 Cancel", callback_data=f"prod_edit|{p_name}"))
        bot.send_message(ADMIN_ID, f"🛡️ *{p_name}* — Select new duration:",
                         reply_markup=mk, parse_mode="Markdown")
    else:
        msg = bot.send_message(ADMIN_ID,
            f"⏱️ New duration for *{p_name}* (e.g. `Lifetime`, `12-36 Month`):", parse_mode="Markdown")
        bot.register_next_step_handler(msg, prod_edit_dur_save_text, p_name)

def prod_edit_dur_save_text(message, p_name):
    if message.chat.id != ADMIN_ID: return
    market_data["products"][p_name]["duration"] = message.text.strip(); save_market_data(market_data)
    bot.send_message(ADMIN_ID, f"✅ *{p_name}* duration updated!", parse_mode="Markdown")

@bot.callback_query_handler(func=lambda c: c.data.startswith("prod_setdurkey|"))
def prod_set_vpn_dur_key(call):
    if call.message.chat.id != ADMIN_ID: return
    _, p_name, dur_key = call.data.split("|", 2)
    durs = get_vpn_durations(); dur_name = durs.get(dur_key, {}).get("name", dur_key)
    if p_name in market_data["products"]:
        market_data["products"][p_name]["vpn_dur"]  = dur_key
        market_data["products"][p_name]["duration"] = dur_name
        save_market_data(market_data)
        bot.answer_callback_query(call.id, f"✅ Moved to {dur_name}!")
        bot.edit_message_text(f"✅ *{p_name}* → *{dur_name}*!",
                              ADMIN_ID, call.message.message_id, parse_mode="Markdown")
    else: bot.answer_callback_query(call.id, "❌ Product not found.")

@bot.callback_query_handler(func=lambda c: c.data.startswith("prod_ename|"))
def prod_edit_name_start(call):
    if call.message.chat.id != ADMIN_ID: return
    p_name = call.data.split("|", 1)[1]; bot.answer_callback_query(call.id)
    msg = bot.send_message(ADMIN_ID, f"✏️ New name for *{p_name}*:", parse_mode="Markdown")
    bot.register_next_step_handler(msg, prod_edit_name_save, p_name)

def prod_edit_name_save(message, p_name):
    if message.chat.id != ADMIN_ID: return
    new_name = message.text.strip()
    if p_name in market_data["products"]:
        market_data["products"][new_name] = market_data["products"].pop(p_name); save_market_data(market_data)
        bot.send_message(ADMIN_ID, f"✅ Renamed to *{new_name}*!", parse_mode="Markdown")

@bot.callback_query_handler(func=lambda c: c.data.startswith("prod_del|"))
def prod_delete(call):
    if call.message.chat.id != ADMIN_ID: return
    p_name = call.data.split("|", 1)[1]
    if p_name in market_data["products"]:
        del market_data["products"][p_name]; save_market_data(market_data)
        bot.answer_callback_query(call.id, f"✅ {p_name} deleted!")
        bot.edit_message_text(f"🗑️ *{p_name}* deleted.",
                              ADMIN_ID, call.message.message_id, parse_mode="Markdown")
    else: bot.answer_callback_query(call.id, "❌ Not found.")


# ─────────────────────────────────────────────
#  Stock Manager (edit / delete / replace stock)
# ─────────────────────────────────────────────
def _adm_stockmgr(call):
    """প্রোডাক্ট সিলেক্ট করার স্ক্রিন।"""
    prods = get_products()
    mk    = types.InlineKeyboardMarkup(row_width=2)
    btns  = [
        types.InlineKeyboardButton(
            f"{p} ({get_stock_count(p)})",
            callback_data=f"smgr|sel|{p}"
        ) for p in prods
    ]
    if btns: mk.add(*btns)
    mk.add(types.InlineKeyboardButton("🔙 Back", callback_data="adm|back"))
    txt = (
        "✏️ *Stock Manager*\n"
        "✨━━━━━━━━━━━━━━━━━━✨\n"
        "কোন প্রোডাক্টের স্টক এডিট করতে চান?"
    )
    try:
        bot.edit_message_text(txt, ADMIN_ID, call.message.message_id,
                              reply_markup=mk, parse_mode="Markdown")
    except Exception:
        bot.send_message(ADMIN_ID, txt, reply_markup=mk, parse_mode="Markdown")


@bot.callback_query_handler(func=lambda c: c.data.startswith("smgr|"))
def stockmgr_router(call):
    if call.message.chat.id != ADMIN_ID: return
    parts  = call.data.split("|", 2)
    action = parts[1]
    p_name = parts[2] if len(parts) > 2 else ""
    bot.answer_callback_query(call.id)

    def _options_markup(p):
        cnt = get_stock_count(p)
        mk2 = types.InlineKeyboardMarkup(row_width=1)
        mk2.add(
            types.InlineKeyboardButton(
                f"📥 স্টক ডাউনলোড করুন ({cnt} পিস)",
                callback_data=f"smgr|dl|{p}"),
            types.InlineKeyboardButton(
                "🔁 স্টক রিপ্লেস করুন (xlsx আপলোড করুন)",
                callback_data=f"smgr|replace|{p}"),
            types.InlineKeyboardButton(
                "🗑️ পুরো স্টক মুছে দিন",
                callback_data=f"smgr|delconfirm|{p}"),
            types.InlineKeyboardButton("🔙 Back", callback_data="adm|stockmgr"),
        )
        return mk2

    if action == "sel":
        cnt = get_stock_count(p_name)
        txt = (
            f"✏️ *{p_name}*\n"
            f"✨━━━━━━━━━━━━━━━━━━✨\n"
            f"📦 বর্তমান স্টক: *{cnt} পিস*\n"
            f"✨━━━━━━━━━━━━━━━━━━✨\n"
            "নিচের অপশনগুলো থেকে বেছে নিন:"
        )
        try:
            bot.edit_message_text(txt, ADMIN_ID, call.message.message_id,
                                  reply_markup=_options_markup(p_name), parse_mode="Markdown")
        except Exception:
            bot.send_message(ADMIN_ID, txt, reply_markup=_options_markup(p_name), parse_mode="Markdown")

    elif action == "dl":
        rows = _load_stock_rows(p_name)
        if not rows:
            bot.send_message(ADMIN_ID, f"❌ *{p_name}* এর কোনো স্টক নেই।", parse_mode="Markdown"); return
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            pd.DataFrame(rows).to_excel(writer, index=False)
        out.seek(0)
        bot.send_document(
            ADMIN_ID, out,
            visible_file_name=f"{p_name}_stock.xlsx",
            caption=(
                f"📥 *{p_name}* — বর্তমান স্টক ({len(rows)} পিস)\n"
                f"✏️ ফাইলটি এডিট করে *রিপ্লেস* অপশনে আপলোড করুন।"
            ),
            parse_mode="Markdown"
        )

    elif action == "replace":
        msg = bot.send_message(
            ADMIN_ID,
            f"📁 *{p_name}* — এডিট করা নতুন Excel (.xlsx) ফাইল পাঠান।\n"
            f"⚠️ এটি পুরনো সমস্ত স্টক বদলে দেবে।",
            parse_mode="Markdown"
        )
        bot.register_next_step_handler(msg, _smgr_do_replace, p_name)

    elif action == "delconfirm":
        mk2 = types.InlineKeyboardMarkup(row_width=2)
        mk2.add(
            types.InlineKeyboardButton("✅ হ্যাঁ, মুছে দিন", callback_data=f"smgr|deldo|{p_name}"),
            types.InlineKeyboardButton("❌ না, বাতিল", callback_data=f"smgr|sel|{p_name}"),
        )
        cnt = get_stock_count(p_name)
        txt = (
            f"⚠️ *নিশ্চিত করুন*\n"
            f"✨━━━━━━━━━━━━━━━━━━✨\n"
            f"📦 *{p_name}* এর সমস্ত *{cnt} পিস* স্টক মুছে দিতে চান?\n"
            f"এই কাজটি পূর্বাবস্থায় ফেরানো যাবে না!"
        )
        try:
            bot.edit_message_text(txt, ADMIN_ID, call.message.message_id,
                                  reply_markup=mk2, parse_mode="Markdown")
        except Exception:
            bot.send_message(ADMIN_ID, txt, reply_markup=mk2, parse_mode="Markdown")

    elif action == "deldo":
        db_delete_stock(p_name)
        bot.send_message(
            ADMIN_ID,
            f"🗑️ *{p_name}* এর সমস্ত স্টক মুছে দেওয়া হয়েছে।",
            parse_mode="Markdown"
        )


def _smgr_do_replace(message, p_name):
    """ইউজার নতুন xlsx পাঠালে পুরনো স্টক বদলে দাও।"""
    if message.chat.id != ADMIN_ID: return
    if not message.document:
        bot.send_message(ADMIN_ID, "❌ Excel (.xlsx) ফাইল পাঠান।"); return
    file_info  = bot.get_file(message.document.file_id)
    downloaded = bot.download_file(file_info.file_path)
    try:
        df   = pd.read_excel(io.BytesIO(downloaded), header=0)
        rows = df.to_dict("records")
    except Exception:
        bot.send_message(ADMIN_ID, "❌ ফাইল পড়া যায়নি। সঠিক Excel (.xlsx) ফাইল পাঠান।"); return
    if not rows:
        bot.send_message(ADMIN_ID, "❌ ফাইলে কোনো ডেটা নেই।"); return
    db_save_stock(p_name, rows)
    bot.send_message(
        ADMIN_ID,
        f"✅ *{p_name}* এর স্টক সফলভাবে রিপ্লেস হয়েছে!\n"
        f"📦 নতুন স্টক: *{len(rows)} পিস*",
        parse_mode="Markdown"
    )


# ─────────────────────────────────────────────
#  Add Stock
# ─────────────────────────────────────────────
def _adm_addstock_select(message):
    prods = get_products()
    mk    = types.InlineKeyboardMarkup(row_width=2)
    btns  = [types.InlineKeyboardButton(p, callback_data=f"as|{p}") for p in prods]
    if btns: mk.add(*btns)
    bot.send_message(ADMIN_ID, "📦 কোন প্রোডাক্টে স্টক আপলোড করতে চান?", reply_markup=mk)

@bot.message_handler(commands=["addstock"])
def admin_add_stock(message):
    if not admin_only(message): return
    _adm_addstock_select(message)

@bot.callback_query_handler(func=lambda c: c.data.startswith("as|"))
def ask_xlsx(call):
    if call.message.chat.id != ADMIN_ID: return
    p_name = call.data.split("|", 1)[1]
    msg = bot.send_message(ADMIN_ID, f"📁 *{p_name}* এর Excel (.xlsx) ফাইল পাঠান:", parse_mode="Markdown")
    bot.register_next_step_handler(msg, save_xlsx, p_name)
    bot.answer_callback_query(call.id)

def save_xlsx(message, p_name):
    if message.chat.id != ADMIN_ID: return
    if not message.document:
        bot.send_message(ADMIN_ID, "❌ Excel (.xlsx) ফাইল পাঠান।"); return
    file_info  = bot.get_file(message.document.file_id)
    downloaded = bot.download_file(file_info.file_path)
    try:
        df   = pd.read_excel(io.BytesIO(downloaded), header=0)
        rows = df.to_dict("records")
    except Exception:
        bot.send_message(ADMIN_ID, "❌ ফাইল পড়া যায়নি। সঠিক Excel (.xlsx) ফাইল পাঠান।"); return
    if not rows:
        bot.send_message(ADMIN_ID, "❌ ফাইলে কোনো ডেটা নেই।"); return
    # Add to MongoDB — merged with any existing unsold stock (survives Railway restarts)
    before_count = get_stock_count(p_name)
    db_append_stock(p_name, rows)
    # Clear synced_stock so we read the real MongoDB-backed count
    if "synced_stock" in market_data["products"].get(p_name, {}):
        del market_data["products"][p_name]["synced_stock"]; save_market_data(market_data)
    stock   = get_stock_count(p_name)
    added   = stock - before_count
    bot.send_message(
        ADMIN_ID,
        f"✅ *{p_name}* স্টক আপডেট হয়েছে!\n"
        f"➕ নতুন যোগ হলো: *{added}* পিস\n"
        f"📦 আগে ছিল: *{before_count}* পিস\n"
        f"📊 বর্তমান মোট স্টক: *{stock}* পিস",
        parse_mode="Markdown",
    )
    # Notify all users about restocked product
    notif = (
        f"📦 *স্টক আপডেট নোটিফিকেশন!*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"🎉 *{p_name}* প্রোডাক্টের স্টক আপডেট হয়েছে!\n"
        f"📊 বর্তমান স্টক: *{stock} পিস*\n"
        f"✨━━━━━━━━━━━━━━━━━━✨\n"
        f"🛒 *এখনই কিনুন, স্টক সীমিত!*"
    )
    for uid in list(user_data.keys()):
        try: bot.send_message(uid, notif, parse_mode="Markdown")
        except Exception: pass


# ─────────────────────────────────────────────
#  Broadcast & Backup & Restore
# ─────────────────────────────────────────────
@bot.message_handler(commands=["broadcast"])
def broadcast_start(message):
    if not admin_only(message): return
    msg = bot.send_message(ADMIN_ID, "📢 সব ইউজারকে কোন মেসেজটি পাঠাতে চান? লিখুন:")
    bot.register_next_step_handler(msg, do_broadcast)

def do_broadcast(message):
    if message.chat.id != ADMIN_ID: return
    text = message.text; success = failed = 0
    for uid in list(user_data.keys()):
        try:
            bot.send_message(uid,
                f"📢 *Admin Broadcast:*\n✨━━━━━━━━━━━━✨\n{text}\n✨━━━━━━━━━━━━✨",
                parse_mode="Markdown")
            success += 1
        except Exception: failed += 1
    bot.send_message(ADMIN_ID, f"✅ Broadcast done!\n📨 Sent: {success}\n❌ Failed: {failed}")

def do_restore(message):
    """Restore user data from a backed-up .json file straight into MongoDB
    (data now lives in MongoDB, not a local file)."""
    if message.chat.id != ADMIN_ID or not message.document: return
    f_info = bot.get_file(message.document.file_id)
    try:
        raw     = bot.download_file(f_info.file_path)
        restored = json.loads(raw.decode("utf-8"))
        db_save_all_users(restored)
        global user_data; user_data = load_data()
        bot.send_message(ADMIN_ID, f"✅ রিস্টোর সফল! ({len(user_data)} জন ইউজার)")
    except Exception as e:
        bot.send_message(ADMIN_ID, f"❌ রিস্টোর ব্যর্থ: {e}")

@bot.message_handler(commands=["restore"])
def restore_db(message):
    if not admin_only(message): return
    msg = bot.send_message(ADMIN_ID, "📥 ইউজার ডাটা (.json) ফাইলটি পাঠান।")
    bot.register_next_step_handler(msg, do_restore)

def _do_backup():
    """Export the live MongoDB-backed data as .json files and send them to
    the admin — everything now lives in MongoDB, this is just a snapshot."""
    now = bst_now()
    bot.send_message(ADMIN_ID, f"💾 *Auto Backup*\n📅 {now} BST", parse_mode="Markdown")
    for data, caption, fname in [
        (user_data,   "📂 User Data",   DATA_FILE),
        (market_data, "📦 Market Data", MARKET_FILE),
        (settings,    "⚙️ Settings",    SETTINGS_FILE),
        (STRINGS,     "📝 Texts",       TEXTS_FILE),
        (coupons,     "🎟️ Coupons",     COUPONS_FILE),
    ]:
        try:
            buf = io.BytesIO(json.dumps(data, indent=4, ensure_ascii=False).encode("utf-8"))
            buf.name = fname
            bot.send_document(ADMIN_ID, buf, visible_file_name=fname, caption=caption)
        except Exception: pass

def _schedule_auto_backup():
    _do_backup()
    t = threading.Timer(86400, _schedule_auto_backup); t.daemon = True; t.start()

@bot.message_handler(commands=["backup"])
def backup_db(message):
    if not admin_only(message): return
    _do_backup()


# ─────────────────────────────────────────────
#  /generate_vps — Ready-to-upload VPS file
# ─────────────────────────────────────────────
@bot.message_handler(commands=["generate_vps"])
def cmd_generate_vps(message):
    """Generate a bot_vps.py with all secrets already filled in and send to admin."""
    if not admin_only(message): return

    bot.send_message(ADMIN_ID,
        "⏳ *VPS ফাইল তৈরি হচ্ছে...*",
        parse_mode="Markdown")
    try:
        # Read the template bot_vps.py from disk
        import os as _os
        vps_path = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "bot_vps.py")
        with open(vps_path, "r", encoding="utf-8") as _f:
            content = _f.read()

        # Collect real secrets from environment
        _bt  = _os.environ.get("BOT_TOKEN",           "") or API_TOKEN
        _aid = _os.environ.get("ADMIN_ID",            str(ADMIN_ID))
        _uid = _os.environ.get("USER_API_ID",         "") or str(USER_API_ID or "")
        _uhash = _os.environ.get("USER_API_HASH",     "") or str(USER_API_HASH or "")
        _sess  = _os.environ.get("USER_SESSION_STRING", "") or str(USER_SESSION or "")
        _mongo = _os.environ.get("MONGODB_URI",       "") or _os.environ.get("MONGO_URI", "") or MONGODB_URI

        # Replace placeholders in CONFIG section
        replacements = [
            ('os.environ.get("BOT_TOKEN",           "YOUR_BOT_TOKEN_HERE")',
             f'os.environ.get("BOT_TOKEN",           "{_bt}")'),
            ('int(os.environ.get("ADMIN_ID",        "7522357347"))',
             f'int(os.environ.get("ADMIN_ID",        "{_aid}"))'),
            ('os.environ.get("USER_API_ID",         "YOUR_API_ID_HERE")',
             f'os.environ.get("USER_API_ID",         "{_uid}")'),
            ('os.environ.get("USER_API_HASH",       "YOUR_API_HASH_HERE")',
             f'os.environ.get("USER_API_HASH",       "{_uhash}")'),
            ('os.environ.get("USER_SESSION_STRING", "YOUR_SESSION_STRING_HERE")',
             f'os.environ.get("USER_SESSION_STRING", "{_sess}")'),
            ('"YOUR_MONGODB_URI_HERE"',
             f'"{_mongo}"'),
        ]
        for old, new in replacements:
            content = content.replace(old, new, 1)

        # Verify placeholders are gone
        remaining = [k for k in ["YOUR_BOT_TOKEN_HERE", "YOUR_API_ID_HERE",
                                  "YOUR_API_HASH_HERE", "YOUR_SESSION_STRING_HERE",
                                  "YOUR_MONGODB_URI_HERE"] if k in content]
        if remaining:
            bot.send_message(ADMIN_ID,
                f"⚠️ কিছু values খুঁজে পাওয়া যায়নি:\n`{remaining}`\n"
                f"Replit Secrets ঠিকমতো set আছে কি?",
                parse_mode="Markdown")
            return

        # Send as file named main.py (ready to upload directly to PyHost)
        import io as _io
        file_bytes = _io.BytesIO(content.encode("utf-8"))
        file_bytes.name = "main.py"
        bot.send_document(
            ADMIN_ID,
            file_bytes,
            caption=(
                "✅ *VPS ফাইল তৈরি হয়েছে!*\n"
                "✨━━━━━━━━━━━━━━━━━━✨\n"
                "📋 *এই ফাইলে সব secrets ভরা আছে।*\n\n"
                "📤 *PyHost-এ Upload করুন:*\n"
                "1️⃣ এই `main.py` ফাইলটি download করুন\n"
                "2️⃣ PyHost Cloud → Upload File → এই ফাইল দিন\n"
                "3️⃣ Deploy & Launch চাপুন\n\n"
                "⚠️ *ফাইলটি কারো সাথে share করবেন না — এতে সব secrets আছে!*\n"
                "✨━━━━━━━━━━━━━━━━━━✨"
            ),
            parse_mode="Markdown"
        )
    except FileNotFoundError:
        bot.send_message(ADMIN_ID,
            "❌ `bot_vps.py` ফাইল খুঁজে পাওয়া যায়নি। Replit-এ ফাইলটি আছে কি?",
            parse_mode="Markdown")
    except Exception as _e:
        bot.send_message(ADMIN_ID,
            f"❌ Error: `{str(_e)[:300]}`",
            parse_mode="Markdown")


# ═══════════════════════════════════════════════
#  MAIN ENTRY POINT
# ═══════════════════════════════════════════════
if __name__ == "__main__":
    keep_alive()

    # Start Pyrogram userbot in background thread
    pyro_thread = threading.Thread(target=_start_pyro_thread, daemon=True)
    pyro_thread.start()

    # Auto-backup after 24h
    t = threading.Timer(86400, _schedule_auto_backup); t.daemon = True; t.start()

    # Clear all user-facing commands — restore ☰ three-line menu
    bot.set_my_commands(
        [
            types.BotCommand("start",    "🚀 Start - বট শুরু করুন"),
            types.BotCommand("get_code", "☎️ Get Code - Read OTP"),
        ],
        scope=types.BotCommandScopeDefault()
    )
    # Set "Get Mail Code" WebApp as the ☰ three-line menu button for all users
    try:
        otp_url = f"https://{_REPLIT_DOMAIN}/code-viewer" if _REPLIT_DOMAIN else ""
        if otp_url:
            bot.set_chat_menu_button(menu_button=types.MenuButtonWebApp(
                type="web_app",
                text="Get Code",
                web_app=types.WebAppInfo(url=otp_url)))
        else:
            bot.set_chat_menu_button(menu_button=types.MenuButtonCommands())
    except Exception as _e:
        print(f"⚠️ MenuButton set: {_e}")
    # Admin keeps their commands + default menu button
    admin_commands = [
        types.BotCommand("admin",       "🔑 Admin Panel"),
        types.BotCommand("manage_user", "👥 Manage a user by ID"),
        types.BotCommand("stats",       "📊 Bot statistics"),
        types.BotCommand("mybalance",   "💰 Supplier bot live balance"),
        types.BotCommand("broadcast",   "📢 Broadcast to all users"),
        types.BotCommand("addstock",    "📦 Upload stock file"),
        types.BotCommand("mailstock",   "📧 View mail stock (Fresh/Used)"),
        types.BotCommand("backup",        "💾 Backup database"),
        types.BotCommand("restore",       "📥 Restore database"),
        types.BotCommand("generate_vps",  "📦 VPS ফাইল generate করুন"),
    ]
    bot.set_my_commands(admin_commands, scope=types.BotCommandScopeChat(chat_id=ADMIN_ID))
    try:
        bot.set_chat_menu_button(chat_id=ADMIN_ID, menu_button=types.MenuButtonCommands())
    except Exception: pass

    print("🚀 Prime Bazar Bot is starting...")

    # Weekly auto-report background thread
    import threading as _threading
    _wr_thread = _threading.Thread(target=_weekly_report_thread, daemon=True)
    _wr_thread.start()

    # Drop any pending webhook (resolves 409 conflicts from old instances)
    try:
        bot.delete_webhook(drop_pending_updates=True)
        print("✅ Webhook cleared / conflict resolved.")
    except Exception as _e:
        print(f"⚠️ Webhook clear: {_e}")

    # Retry loop — if 409 conflict persists, wait and retry
    while True:
        try:
            bot.infinity_polling(
                timeout=60,
                long_polling_timeout=30,
                restart_on_change=False,
                allowed_updates=["message", "callback_query"],
            )
        except Exception as _poll_err:
            err_str = str(_poll_err)
            if "409" in err_str:
                print("⚠️ 409 Conflict — another instance running. Retrying in 15s...")
                time.sleep(15)
                try:
                    bot.delete_webhook(drop_pending_updates=True)
                except Exception:
                    pass
            else:
                print(f"⚠️ Polling error: {_poll_err}. Retrying in 5s...")
                time.sleep(5)
