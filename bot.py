import os, io, logging, asyncio, sqlite3
import openpyxl, xlrd
from dotenv import load_dotenv
from telegram import Update, ReplyKeyboardMarkup, InlineKeyboardMarkup, InlineKeyboardButton
from telegram.ext import Application, MessageHandler, CallbackQueryHandler, ContextTypes, filters

load_dotenv()
BOT_TOKEN     = os.environ["TELEGRAM_BOT_TOKEN"]
GROUP_CHAT_ID = os.environ["TELEGRAM_GROUP_CHAT_ID"]
ADMIN_IDS     = [int(x.strip()) for x in os.environ.get("ADMIN_IDS","").split(",") if x.strip()]
DB_PATH       = os.path.expanduser("~/tg-bot/data.db")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
METHOD_LABELS = {"bkash":"bKash","nagad":"Nagad","binance":"Binance UID"}
METHOD_EMOJI  = {"bkash":"🟣","nagad":"🟠","binance":"🟡"}
MAIN_KB = ReplyKeyboardMarkup([["💰 ব্যালেন্স","📁 ফাইল সাবমিট"],["💸 উইথড্র","🆘 সাপোর্ট"]],resize_keyboard=True,is_persistent=True)
CANCEL_KB = ReplyKeyboardMarkup([["❌ বাতিল করুন"]],resize_keyboard=True)
user_states: dict[int,dict] = {}

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_conn(); cur = conn.cursor()
    cur.executescript("""
        CREATE TABLE IF NOT EXISTS tg_users(telegram_id TEXT PRIMARY KEY,username TEXT,first_name TEXT NOT NULL,balance REAL NOT NULL DEFAULT 0,created_at TEXT NOT NULL DEFAULT(datetime('now')));
        CREATE TABLE IF NOT EXISTS submissions(id INTEGER PRIMARY KEY AUTOINCREMENT,telegram_id TEXT NOT NULL,file_id TEXT NOT NULL,file_unique_id TEXT NOT NULL,file_name TEXT NOT NULL,row_count INTEGER NOT NULL,status TEXT NOT NULL DEFAULT 'pending',admin_note TEXT,group_message_id INTEGER,created_at TEXT NOT NULL DEFAULT(datetime('now')),updated_at TEXT NOT NULL DEFAULT(datetime('now')));
        CREATE TABLE IF NOT EXISTS withdrawals(id INTEGER PRIMARY KEY AUTOINCREMENT,telegram_id TEXT NOT NULL,amount REAL NOT NULL,method TEXT NOT NULL DEFAULT 'bkash',account_number TEXT NOT NULL DEFAULT '',status TEXT NOT NULL DEFAULT 'pending',created_at TEXT NOT NULL DEFAULT(datetime('now')));
    """)
    conn.commit(); conn.close()

def _q(sql, params=None, fetch="one"):
    conn = get_conn(); cur = conn.cursor()
    cur.execute(sql, params or [])
    if fetch == "one":
        row = cur.fetchone()
        result = dict(row) if row else None
    elif fetch == "all":
        rows = cur.fetchall()
        result = [dict(r) for r in rows]
    else:
        result = None
    conn.commit(); conn.close()
    return result

def db_get_user(tid): return _q("SELECT * FROM tg_users WHERE telegram_id=?",[str(tid)])
def db_get_or_create(tid,fname,uname=None):
    row=db_get_user(tid)
    if not row:
        _q("INSERT OR IGNORE INTO tg_users(telegram_id,first_name,username) VALUES(?,?,?)",[str(tid),fname,uname],fetch="none")
        row=db_get_user(tid)
    return row
def db_set_balance(tid,bal): _q("UPDATE tg_users SET balance=? WHERE telegram_id=?",[bal,str(tid)],fetch="none")
def db_insert_sub(tid,fid,fuid,fname,rc):
    _q("INSERT INTO submissions(telegram_id,file_id,file_unique_id,file_name,row_count) VALUES(?,?,?,?,?)",[str(tid),fid,fuid,fname,rc],fetch="none")
    return _q("SELECT * FROM submissions WHERE telegram_id=? ORDER BY id DESC LIMIT 1",[str(tid)])
def db_set_sub_group_msg(sid,gid): _q("UPDATE submissions SET group_message_id=? WHERE id=?",[gid,sid],fetch="none")
def db_get_sub(sid): return _q("SELECT * FROM submissions WHERE id=?",[sid])
def db_set_sub_status(sid,st): _q("UPDATE submissions SET status=?,updated_at=datetime('now') WHERE id=?",[st,sid],fetch="none")
def db_insert_wd(tid,amt,method,acc):
    _q("INSERT INTO withdrawals(telegram_id,amount,method,account_number) VALUES(?,?,?,?)",[str(tid),amt,method,acc],fetch="none")
    return _q("SELECT * FROM withdrawals WHERE telegram_id=? ORDER BY id DESC LIMIT 1",[str(tid)])
def db_get_wd(wid): return _q("SELECT * FROM withdrawals WHERE id=?",[wid])
def db_set_wd_status(wid,st): _q("UPDATE withdrawals SET status=? WHERE id=?",[st,wid],fetch="none")
def db_list_users(n=20): return _q("SELECT * FROM tg_users ORDER BY created_at DESC LIMIT ?",[n],fetch="all")
def db_list_pending_subs(n=10): return _q("SELECT * FROM submissions WHERE status='pending' ORDER BY created_at DESC LIMIT ?",[n],fetch="all")
def db_list_pending_wds(n=10): return _q("SELECT * FROM withdrawals WHERE status='pending' ORDER BY created_at DESC LIMIT ?",[n],fetch="all")

def count_col_a(fbytes,fname):
    try:
        if fname.lower().endswith(".xlsx"):
            wb=openpyxl.load_workbook(io.BytesIO(fbytes),read_only=True,data_only=True)
            ws=wb.active; c=sum(1 for r in ws.iter_rows(min_col=1,max_col=1,values_only=True) if r[0] is not None and str(r[0]).strip()); wb.close(); return c
        elif fname.lower().endswith(".xls"):
            wb=xlrd.open_workbook(file_contents=fbytes); ws=wb.sheet_by_index(0)
            return sum(1 for i in range(ws.nrows) if ws.cell_value(i,0) and str(ws.cell_value(i,0)).strip())
    except Exception as e: logger.error(f"Excel error:{e}")
    return 0

def is_excel(name,mime=""):
    n=name.lower(); return n.endswith(".xlsx") or n.endswith(".xls") or mime in("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","application/vnd.ms-excel")
def is_admin(uid): return uid in ADMIN_IDS

async def handle_message(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    msg=update.message; user=msg.from_user
    if not user: return
    uid=user.id; text=msg.text or ""; state=user_states.get(uid)
    if text=="❌ বাতিল করুন":
        user_states.pop(uid,None); await msg.reply_text("✅ বাতিল করা হয়েছে।",reply_markup=MAIN_KB); return
    if state:
        s=state["step"]
        if s=="WAITING_FILE":             await do_file_upload(update,ctx); return
        if s=="WAITING_WITHDRAW_ACCOUNT": await do_wd_account(update,ctx,state); return
        if s=="WAITING_WITHDRAW_AMOUNT":  await do_wd_amount(update,ctx,state); return
        if s=="WAITING_SUPPORT_MSG":      await do_support_msg(update,ctx); return
        if s=="WAITING_APPROVE_COUNT":    await do_approve_count(update,ctx,state); return
        if s=="WAITING_APPROVE_RATE":     await do_approve_rate(update,ctx,state); return
    if is_admin(uid) and text.startswith("/") and text!=("/start"):
        await do_admin_cmd(update,ctx); return
    menu={"/start":do_start,"💰 ব্যালেন্স":do_balance,"📁 ফাইল সাবমিট":do_submit_file,"💸 উইথড্র":do_withdraw,"🆘 সাপোর্ট":do_support}
    fn=menu.get(text)
    if fn: await fn(update,ctx)
    else: await msg.reply_text("নিচের বাটন ব্যবহার করুন।",reply_markup=MAIN_KB)

async def do_start(update,ctx):
    u=update.message.from_user; db_get_or_create(u.id,u.first_name,u.username)
    await update.message.reply_text(f"স্বাগতম, {u.first_name}! 👋\n\nনিচের বাটন থেকে আপনার কাজ করুন।",reply_markup=MAIN_KB)

async def do_balance(update,ctx):
    u=update.message.from_user; row=db_get_or_create(u.id,u.first_name,u.username)
    await update.message.reply_text(f"💰 *আপনার ব্যালেন্স:* `{float(row['balance']):.2f} টাকা`",parse_mode="Markdown",reply_markup=MAIN_KB)

async def do_submit_file(update,ctx):
    user_states[update.message.from_user.id]={"step":"WAITING_FILE"}
    await update.message.reply_text("📁 Excel ফাইল পাঠান (.xlsx বা .xls)\n\nবাতিল করতে নিচের বাটন চাপুন।",reply_markup=CANCEL_KB)

async def do_file_upload(update,ctx):
    msg=update.message; user=msg.from_user; uid=user.id; doc=msg.document
    if not doc: await msg.reply_text("❌ শুধুমাত্র Excel ফাইল পাঠান।"); return
    fname=doc.file_name or "file"
    if not is_excel(fname,doc.mime_type or ""):
        await msg.reply_text("❌ শুধুমাত্র Excel ফাইল (.xlsx বা .xls) গ্রহণযোগ্য।"); return
    user_states.pop(uid,None)
    proc=await msg.reply_text("⏳ ফাইল প্রক্রিয়া হচ্ছে...")
    tgf=await ctx.bot.get_file(doc.file_id); fbytes=await tgf.download_as_bytearray()
    rc=count_col_a(bytes(fbytes),fname)
    db_get_or_create(uid,user.first_name,user.username)
    sub=db_insert_sub(uid,doc.file_id,doc.file_unique_id,fname,rc)
    un=f" (@{user.username})" if user.username else ""
    cap=(f"📁 *নতুন ফাইল সাবমিশন*\n\n👤 ইউজার: [{user.first_name}](tg://user?id={uid}){un}\n🆔 ID: `{uid}`\n📄 ফাইল: {fname}\n📊 মোট আইডি: *{rc} টি*\n🔢 সাবমিশন ID: `{sub['id']}`")
    grp=await ctx.bot.send_document(chat_id=GROUP_CHAT_ID,document=doc.file_id,caption=cap,parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("✅ অ্যাপ্রুভ",callback_data=f"approve_{sub['id']}"),InlineKeyboardButton("❌ রিজেক্ট",callback_data=f"reject_{sub['id']}")]])) 
    db_set_sub_group_msg(sub["id"],grp.message_id)
    await proc.delete()
    await msg.reply_text(f"✅ *ফাইল সাবমিট হয়েছে!*\n\n📄 ফাইল: {fname}\n📊 মোট আইডি: *{rc} টি*\n\nঅ্যাপ্রুভ বা রিজেক্টের জন্য অপেক্ষা করুন।",parse_mode="Markdown",reply_markup=MAIN_KB)

async def do_withdraw(update,ctx):
    u=update.message.from_user; row=db_get_or_create(u.id,u.first_name,u.username); bal=float(row["balance"])
    if bal<=0: await update.message.reply_text(f"❌ আপনার ব্যালেন্স নেই।\n\nবর্তমান ব্যালেন্স: *{bal:.2f} টাকা*",parse_mode="Markdown",reply_markup=MAIN_KB); return
    await update.message.reply_text(f"💸 *উইথড্র পদ্ধতি বেছে নিন*\n\nবর্তমান ব্যালেন্স: *{bal:.2f} টাকা*",parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🟣 bKash",callback_data="wd_method_bkash")],[InlineKeyboardButton("🟠 Nagad",callback_data="wd_method_nagad")],[InlineKeyboardButton("🟡 Binance UID",callback_data="wd_method_binance")]]))

async def do_wd_account(update,ctx,state):
    msg=update.message; uid=msg.from_user.id; account=(msg.text or "").strip()
    if not account: await msg.reply_text("❌ সঠিক নম্বর/UID লিখুন।"); return
    method=state["method"]; ml=state["method_label"]; fee=5 if method in("bkash","nagad") else 0
    fn=f"\n⚠️ {ml}-এ *{fee} টাকা* সার্ভিস ফি কাটবে।" if fee else ""
    user_states[uid]={"step":"WAITING_WITHDRAW_AMOUNT","method":method,"method_label":ml,"account":account}
    row=db_get_user(uid); bal=float(row["balance"]) if row else 0
    await msg.reply_text(f"✅ {ml} নম্বর: `{account}`{fn}\n\nকত টাকা উইথড্র করবেন?\n_(ন্যূনতম: ২০ টাকা | ব্যালেন্স: {bal:.2f} টাকা)_",parse_mode="Markdown")

async def do_wd_amount(update,ctx,state):
    msg=update.message; uid=msg.from_user.id
    try: amount=float((msg.text or "").replace(",",""))
    except: await msg.reply_text("❌ সঠিক পরিমাণ লিখুন।"); return
    if amount<20: await msg.reply_text("❌ ন্যূনতম উইথড্র *২০ টাকা*।",parse_mode="Markdown"); return
    method=state["method"]; ml=state["method_label"]; acc=state["account"]; fee=5 if method in("bkash","nagad") else 0
    total=amount+fee; row=db_get_user(uid)
    if not row: return
    bal=float(row["balance"])
    if total>bal:
        ft=f" + *{fee} টাকা ফি*" if fee else ""
        await msg.reply_text(f"❌ পর্যাপ্ত ব্যালেন্স নেই।\n\nঅনুরোধ: *{amount:.2f} টাকা*{ft}\nমোট: *{total:.2f} টাকা*\nব্যালেন্স: *{bal:.2f} টাকা*",parse_mode="Markdown"); return
    user_states.pop(uid,None); new_bal=bal-total; db_set_balance(uid,new_bal)
    wd=db_insert_wd(uid,amount,method,acc)
    emoji=METHOD_EMOJI.get(method,"💸"); un=f" (@{msg.from_user.username})" if msg.from_user.username else ""
    fl=f"💳 সার্ভিস ফি: *{fee} টাকা*\n" if fee else ""
    await ctx.bot.send_message(GROUP_CHAT_ID,
        f"💸 *নতুন উইথড্র অনুরোধ*\n\n👤 [{msg.from_user.first_name}](tg://user?id={uid}){un}\n🆔 ID: `{uid}`\n{emoji} {ml}: `{acc}`\n💰 *{amount:.2f} টাকা*\n{fl}💼 বাকি: *{new_bal:.2f} টাকা*\n🔢 ID: `{wd['id']}`",
        parse_mode="Markdown",reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("✅ অ্যাপ্রুভ",callback_data=f"approvew_{wd['id']}"),InlineKeyboardButton("❌ রিজেক্ট",callback_data=f"rejectw_{wd['id']}")]])) 
    fd=f"\n💳 ফি: *{fee} টাকা*" if fee else ""
    await msg.reply_text(f"✅ *উইথড্র অনুরোধ পাঠানো হয়েছে!*\n\n{emoji} {ml}: `{acc}`\n💰 *{amount:.2f} টাকা*{fd}\n\nঅ্যাডমিন রিভিউ করবেন।",parse_mode="Markdown",reply_markup=MAIN_KB)

async def do_support(update,ctx):
    user_states[update.message.from_user.id]={"step":"WAITING_SUPPORT_MSG"}
    await update.message.reply_text("🆘 আপনার সমস্যা বা প্রশ্ন লিখুন।",reply_markup=CANCEL_KB)

async def do_support_msg(update,ctx):
    msg=update.message; u=msg.from_user; uid=u.id
    user_states.pop(uid,None); un=f" (@{u.username})" if u.username else ""
    for aid in ADMIN_IDS:
        try: await ctx.bot.send_message(aid,f"🆘 *সাপোর্ট মেসেজ*\n\n👤 {u.first_name}{un}\n🆔 `{uid}`\n\n📝 {msg.text}\n\nজবাব: /msg {uid} <মেসেজ>",parse_mode="Markdown")
        except: pass
    await msg.reply_text("✅ মেসেজ অ্যাডমিনকে পাঠানো হয়েছে।",reply_markup=MAIN_KB)

async def do_approve_count(update,ctx,state):
    msg=update.message; uid=msg.from_user.id
    try: count=int((msg.text or "").strip())
    except: await msg.reply_text("❌ সঠিক সংখ্যা লিখুন। যেমন: 100"); return
    if count<=0: await msg.reply_text("❌ সঠিক সংখ্যা লিখুন।"); return
    sub=db_get_sub(state["subId"])
    if not sub: user_states.pop(uid,None); await msg.reply_text("❌ সাবমিশন পাওয়া যায়নি।"); return
    user_states[uid]={"step":"WAITING_APPROVE_RATE","subId":state["subId"],"approvedCount":count,"groupChatId":state["groupChatId"],"groupMessageId":state["groupMessageId"],"fileName":sub["file_name"],"telegramId":sub["telegram_id"]}
    await msg.reply_text(f"✅ ID সংখ্যা: *{count} টি*\n\n💰 প্রতি ID রেট কত টাকা?\n_(যেমন: 2.8)_",parse_mode="Markdown")

async def do_approve_rate(update,ctx,state):
    msg=update.message; uid=msg.from_user.id
    try: rate=float((msg.text or "").replace(",",".").strip())
    except: await msg.reply_text("❌ সঠিক রেট লিখুন। যেমন: 2.8"); return
    if rate<=0: await msg.reply_text("❌ সঠিক রেট লিখুন।"); return
    user_states.pop(uid,None)
    count=state["approvedCount"]; earned=round(count*rate,2)
    row=db_get_user(int(state["telegramId"])); cur_bal=float(row["balance"]) if row else 0
    new_bal=round(cur_bal+earned,2); db_set_balance(int(state["telegramId"]),new_bal)
    try:
        await ctx.bot.send_message(int(state["telegramId"]),f"💰 *পেমেন্ট যোগ হয়েছে!*\n\n📄 ফাইল: {state['fileName']}\n📊 আইডি: *{count} টি*\n💵 রেট: *{rate} টাকা/ID*\n💰 মোট আয়: *{earned:.2f} টাকা*\n🏦 বর্তমান ব্যালেন্স: *{new_bal:.2f} টাকা*",parse_mode="Markdown",reply_markup=MAIN_KB)
    except: pass
    try:
        await ctx.bot.edit_message_reply_markup(chat_id=state["groupChatId"],message_id=state["groupMessageId"],
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton(f"✅ পেমেন্ট দেওয়া হয়েছে ({count} ID × {rate}৳)",callback_data="done")]]))
    except: pass
    await msg.reply_text(f"✅ *পেমেন্ট সম্পন্ন!*\n\n📄 {state['fileName']}\n📊 ID: *{count} টি*\n💵 রেট: *{rate} টাকা/ID*\n💰 যোগ হয়েছে: *{earned:.2f} টাকা*",parse_mode="Markdown")

async def handle_callback(update,ctx):
    q=update.callback_query; uid=q.from_user.id; data=q.data or ""
    if data.startswith("wd_method_"):
        method=data.replace("wd_method_",""); ml=METHOD_LABELS.get(method,method)
        row=db_get_user(uid)
        if not row: await q.answer("প্রথমে /start করুন।"); return
        if float(row["balance"])<=0: await q.answer("❌ ব্যালেন্স নেই।"); return
        user_states[uid]={"step":"WAITING_WITHDRAW_ACCOUNT","method":method,"method_label":ml}
        emoji=METHOD_EMOJI.get(method,"💸"); al="Binance UID" if method=="binance" else f"{ml} নম্বর"
        await q.answer(); await q.message.reply_text(f"{emoji} *{ml}* সিলেক্ট।\n\নআপনার {al} লিখুন:",parse_mode="Markdown",reply_markup=CANCEL_KB); return
    if data.startswith("approvew_") or data.startswith("rejectw_"):
        if not is_admin(uid): await q.answer("❌ অ্যাডমিন নন।"); return
        is_app=data.startswith("approvew_"); wid=int(data.split("_")[1]); wd=db_get_wd(wid)
        if not wd: await q.answer("উইথড্র পাওয়া যায়নি।"); return
        if wd["status"]!="pending": await q.answer("আগেই প্রসেস হয়েছে।"); return
        emoji=METHOD_EMOJI.get(wd["method"],"💸"); label=METHOD_LABELS.get(wd["method"],wd["method"])
        if is_app:
            db_set_wd_status(wid,"approved"); urow=db_get_user(int(wd["telegram_id"]))
            try: await ctx.bot.send_message(int(wd["telegram_id"]),f"✅ *উইথড্র অ্যাপ্রুভ হয়েছে!*\n\n{emoji} {label}: `{wd['account_number']}`\n💰 *{float(wd['amount']):.2f} টাকা*\nব্যালেন্স: *{float(urow['balance']):.2f} টাকা*",parse_mode="Markdown")
            except: pass
            await q.edit_message_reply_markup(InlineKeyboardMarkup([[InlineKeyboardButton("✅ অ্যাপ্রুভ হয়েছে",callback_data="done")]])); await q.answer("✅ অ্যাপ্রুভ।")
        else:
            db_set_wd_status(wid,"rejected"); urow=db_get_user(int(wd["telegram_id"]))
            refund=float(urow["balance"])+float(wd["amount"]) if urow else float(wd["amount"]); db_set_balance(int(wd["telegram_id"]),refund)
            try: await ctx.bot.send_message(int(wd["telegram_id"]),f"❌ *উইথড্র রিজেক্ট হয়েছে।*\n\n{emoji} {label}: {float(wd['amount']):.2f} টাকা\n♻️ ব্যালেন্স ফেরত। বর্তমান: *{refund:.2f} টাকা*",parse_mode="Markdown")
            except: pass
            await q.edit_message_reply_markup(InlineKeyboardMarkup([[InlineKeyboardButton("❌ রিজেক্ট হয়েছে",callback_data="done")]])); await q.answer("❌ রিজেক্ট।")
        return
    if data.startswith("approve_") or data.startswith("reject_"):
        if not is_admin(uid): await q.answer("❌ অ্যাডমিন নন।"); return
        is_app=data.startswith("approve_"); sid=int(data.split("_")[1]); sub=db_get_sub(sid)
        if not sub: await q.answer("সাবমিশন পাওয়া যায়নি।"); return
        if sub["status"]!="pending": await q.answer("আগেই প্রসেস হয়েছে।"); return
        if is_app:
            db_set_sub_status(sid,"approved")
            try: await ctx.bot.send_message(int(sub["telegram_id"]),f"✅ *আপনার ফাইল অ্যাপ্রুভ হয়েছে!*\n\n📄 ফাইল: {sub['file_name']}\n📊 মোট আইডি: *{sub['row_count']} টি*\n\n⏳ পেমেন্ট প্রক্রিয়া চলছে, শীঘ্রই ব্যালেন্সে যোগ হবে।",parse_mode="Markdown",reply_markup=MAIN_KB)
            except: pass
            await q.edit_message_reply_markup(InlineKeyboardMarkup([[InlineKeyboardButton("💰 রিপোর্ট দিন",callback_data=f"report_{sid}")]])); await q.answer("✅ অ্যাপ্রুভ। রিপোর্ট দিন।")
        else:
            db_set_sub_status(sid,"rejected")
            try: await ctx.bot.send_message(int(sub["telegram_id"]),f"❌ *আপনার ফাইল রিজেক্ট হয়েছে।*\n\n📄 {sub['file_name']}\n\nসমস্যা থাকলে সাপোর্টে যোগাযোগ করুন।",parse_mode="Markdown",reply_markup=MAIN_KB)
            except: pass
            await q.edit_message_reply_markup(InlineKeyboardMarkup([[InlineKeyboardButton("❌ রিজেক্ট হয়েছে",callback_data="done")]])); await q.answer("❌ রিজেক্ট।")
        return
    if data.startswith("report_"):
        if not is_admin(uid): await q.answer("❌ অ্যাডমিন নন।"); return
        sid=int(data.replace("report_","")); sub=db_get_sub(sid)
        if not sub: await q.answer("সাবমিশন পাওয়া যায়নি।"); return
        user_states[uid]={"step":"WAITING_APPROVE_COUNT","subId":sid,"groupChatId":q.message.chat.id,"groupMessageId":q.message.message_id}
        await q.answer()
        await ctx.bot.send_message(uid,f"📊 *রিপোর্ট দিন*\n\n📄 ফাইল: {sub['file_name']}\n📁 মোট আইডি: *{sub['row_count']} টি*\n\n✏️ কতটি ID রিপোর্ট দিচ্ছেন?",parse_mode="Markdown"); return
    await q.answer()

async def do_admin_cmd(update,ctx):
    msg=update.message; parts=(msg.text or "").split(); cmd=parts[0].lower()
    if cmd=="/addbalance":
        if len(parts)<3: await msg.reply_text("ব্যবহার: /addbalance <id> <amount>"); return
        try: amt=float(parts[2])
        except: await msg.reply_text("সঠিক পরিমাণ দিন।"); return
        row=db_get_user(int(parts[1]))
        if not row: await msg.reply_text("❌ ইউজার পাওয়া যায়নি।"); return
        new=float(row["balance"])+amt; db_set_balance(int(parts[1]),new)
        await msg.reply_text(f"✅ +{amt} টাকা যোগ। নতুন ব্যালেন্স: {new:.2f} টাকা")
        try: await ctx.bot.send_message(int(parts[1]),f"💰 *ব্যালেন্সে যোগ হয়েছে!*\n+{amt} টাকা\nবর্তমান: {new:.2f} টাকা",parse_mode="Markdown")
        except: pass; return
    if cmd=="/deductbalance":
        if len(parts)<3: await msg.reply_text("ব্যবহার: /deductbalance <id> <amount>"); return
        try: amt=float(parts[2])
        except: await msg.reply_text("সঠিক পরিমাণ দিন।"); return
        row=db_get_user(int(parts[1]))
        if not row: await msg.reply_text("❌ ইউজার পাওয়া যায়নি।"); return
        new=max(0.0,float(row["balance"])-amt); db_set_balance(int(parts[1]),new)
        await msg.reply_text(f"✅ -{amt} টাকা কাটা। নতুন ব্যালেন্স: {new:.2f} টাকা"); return
    if cmd=="/msg":
        if len(parts)<3: await msg.reply_text("ব্যবহার: /msg <id> <মেসেজ>"); return
        try: await ctx.bot.send_message(int(parts[1]),f"📨 *অ্যাডমিনের মেসেজ:*\n\n{' '.join(parts[2:])}",parse_mode="Markdown"); await msg.reply_text("✅ মেসেজ পাঠানো হয়েছে।")
        except: await msg.reply_text("❌ মেসেজ পাঠানো যায়নি।")
        return
    if cmd=="/broadcast":
        text=" ".join(parts[1:]).strip()
        if not text: await msg.reply_text("ব্যবহার: /broadcast <মেসেজ>"); return
        users=db_list_users(9999)
        if not users: await msg.reply_text("কোনো ইউজার নেই।"); return
        st=await msg.reply_text(f"📡 পাঠানো হচ্ছে ({len(users)} জন)..."); ok=fail=0
        for u in users:
            try: await ctx.bot.send_message(int(u["telegram_id"]),f"📢 *অ্যাডমিনের বার্তা:*\n\n{text}",parse_mode="Markdown"); ok+=1
            except: fail+=1
            await asyncio.sleep(0.05)
        await st.edit_text(f"✅ *ব্রডকাস্ট সম্পন্ন!*\n\n✅ সফল: {ok} জন\n❌ ব্যর্থ: {fail} জন",parse_mode="Markdown"); return
    if cmd=="/users":
        rows=db_list_users()
        if not rows: await msg.reply_text("কোনো ইউজার নেই।"); return
        lines="\n\n".join(f"{i+1}. {u['first_name']}{' (@'+u['username']+')' if u['username'] else ''}\n   ID: `{u['telegram_id']}` | {float(u['balance']):.2f} টাকা" for i,u in enumerate(rows))
        await msg.reply_text(f"👥 *ইউজার (সর্বশেষ ২০)*\n\n{lines}",parse_mode="Markdown"); return
    if cmd=="/submissions":
        rows=db_list_pending_subs()
        if not rows: await msg.reply_text("কোনো পেন্ডিং সাবমিশন নেই।"); return
        lines="\n\n".join(f"🔢 ID: `{s['id']}` | 👤 `{s['telegram_id']}`\n📄 {s['file_name']} | 📊 {s['row_count']} টি" for s in rows)
        await msg.reply_text(f"📁 *পেন্ডিং সাবমিশন*\n\n{lines}",parse_mode="Markdown"); return
    if cmd=="/withdrawals":
        rows=db_list_pending_wds()
        if not rows: await msg.reply_text("কোনো পেন্ডিং উইথড্র নেই।"); return
        lines="\n\n".join(f"🔢 ID: `{w['id']}` | 👤 `{w['telegram_id']}`\n{METHOD_EMOJI.get(w['method'],'💸')} {METHOD_LABELS.get(w['method'],w['method'])}: `{w['account_number']}`\n💰 *{float(w['amount']):.2f} টাকা*" for w in rows)
        await msg.reply_text(f"💸 *পেন্ডিং উইথড্র*\n\n{lines}",parse_mode="Markdown"); return
    if cmd=="/approvewithdraw":
        if len(parts)<2: await msg.reply_text("ব্যবহার: /approvewithdraw <id>"); return
        wd=db_get_wd(int(parts[1]))
        if not wd: await msg.reply_text("❌ উইথড্র পাওয়া যায়নি।"); return
        if wd["status"]!="pending": await msg.reply_text("আগেই প্রসেস হয়েছে।"); return
        db_set_wd_status(int(parts[1]),"approved")
        emoji=METHOD_EMOJI.get(wd["method"],"💸"); label=METHOD_LABELS.get(wd["method"],wd["method"])
        urow=db_get_user(int(wd["telegram_id"]))
        try: await ctx.bot.send_message(int(wd["telegram_id"]),f"✅ *উইথড্র অ্যাপ্রুভ!*\n\n{emoji} {label}: `{wd['account_number']}`\n💰 *{float(wd['amount']):.2f} টাকা*\nব্যালেন্স: *{float(urow['balance']):.2f} টাকা*",parse_mode="Markdown")
        except: pass
        await msg.reply_text("✅ উইথড্র অ্যাপ্রুভ করা হয়েছে।"); return
    if cmd=="/rejectwithdraw":
        if len(parts)<2: await msg.reply_text("ব্যবহার: /rejectwithdraw <id>"); return
        wd=db_get_wd(int(parts[1]))
        if not wd: await msg.reply_text("❌ উইথড্র পাওয়া যায়নি।"); return
        if wd["status"]!="pending": await msg.reply_text("আগেই প্রসেস হয়েছে।"); return
        db_set_wd_status(int(parts[1]),"rejected")
        urow=db_get_user(int(wd["telegram_id"])); refund=float(urow["balance"])+float(wd["amount"]) if urow else float(wd["amount"])
        db_set_balance(int(wd["telegram_id"]),refund)
        emoji=METHOD_EMOJI.get(wd["method"],"💸"); label=METHOD_LABELS.get(wd["method"],wd["method"])
        try: await ctx.bot.send_message(int(wd["telegram_id"]),f"❌ *উইথড্র রিজেক্ট।*\n\n{emoji} {label}: {float(wd['amount']):.2f} টাকা\n♻️ ব্যালেন্স ফেরত। বর্তমান: *{refund:.2f} টাকা*",parse_mode="Markdown")
        except: pass
        await msg.reply_text("❌ রিজেক্ট। ব্যালেন্স ফেরত দেওয়া হয়েছে।"); return
    if cmd=="/stats":
        tu=_q("SELECT COUNT(*) AS c FROM tg_users")["c"]
        ts=_q("SELECT COUNT(*) AS c FROM submissions")["c"]
        as_=_q("SELECT COUNT(*) AS c FROM submissions WHERE status='approved'")["c"]
        ps=_q("SELECT COUNT(*) AS c FROM submissions WHERE status='pending'")["c"]
        tw=_q("SELECT COALESCE(SUM(amount),0) AS s FROM withdrawals WHERE status='approved'")["s"]
        pw=_q("SELECT COUNT(*) AS c FROM withdrawals WHERE status='pending'")["c"]
        pa=_q("SELECT COALESCE(SUM(amount),0) AS s FROM withdrawals WHERE status='pending'")["s"]
        tb=_q("SELECT COALESCE(SUM(balance),0) AS s FROM tg_users")["s"]
        await msg.reply_text(f"📊 *বট পরিসংখ্যান*\n\n👥 মোট ইউজার: *{tu} জন*\n\n📁 মোট সাবমিশন: *{ts} টি*\n   ✅ অ্যাপ্রুভ: *{as_} টি*\n   ⏳ পেন্ডিং: *{ps} টি*\n\n💸 মোট পেমেন্ট দেওয়া: *{float(tw):.2f} টাকা*\n⏳ পেন্ডিং উইথড্র: *{pw} টি* ({float(pa):.2f} টাকা)\n\n🏦 ইউজারদের মোট ব্যালেন্স: *{float(tb):.2f} টাকা*",parse_mode="Markdown"); return
    if cmd=="/help":
        await msg.reply_text("🛠 *অ্যাডমিন কমান্ড:*\n\n/addbalance <id> <amount>\n/deductbalance <id> <amount>\n/msg <id> <মেসেজ>\n/broadcast <মেসেজ>\n/users\n/submissions\n/withdrawals\n/approvewithdraw <id>\n/rejectwithdraw <id>\n/stats",parse_mode="Markdown")

def main():
    init_db()
    app=Application.builder().token(BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.ALL,handle_message))
    app.add_handler(CallbackQueryHandler(handle_callback))
    logger.info("✅ Bot started! (SQLite)")
    app.run_polling(drop_pending_updates=True)

if __name__=="__main__":
    main()
