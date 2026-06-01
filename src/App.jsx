import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, LineChart, Line, CartesianGrid } from "recharts";

/* ─────────────────────────────────────────
   디자인 토큰 — 전문 회계법인 스타일
   네이비 딥톤 + 골드 포인트 + 절제된 여백
───────────────────────────────────────── */
const T = {
  // 배경
  bg:          "#F0F2F5",
  bgDeep:      "#1B2A4A",       // 네이비 딥 (헤더·사이드바)
  bgDeepAlt:   "#152240",       // 더 진한 네이비
  surface:     "#FFFFFF",
  surfaceNav:  "rgba(27,42,74,0.97)",

  // 테두리
  border:      "rgba(0,0,0,0.09)",
  borderStrong:"rgba(0,0,0,0.16)",
  borderNav:   "rgba(255,255,255,0.1)",

  // 텍스트
  text:        "#1A1F2E",
  textSub:     "#4A5568",
  textMuted:   "#8896A8",
  textNav:     "#CBD5E0",       // 네이비 배경 위 텍스트
  textNavActive:"#FFFFFF",

  // 브랜드 컬러
  gold:        "#C9A84C",       // 골드 포인트
  goldLight:   "rgba(201,168,76,0.12)",
  goldBorder:  "rgba(201,168,76,0.35)",
  navy:        "#1B2A4A",
  navyLight:   "rgba(27,42,74,0.07)",

  // 상태 컬러
  blue:        "#1E5FAD",
  blueLight:   "rgba(30,95,173,0.1)",
  green:       "#2E7D52",
  greenLight:  "rgba(46,125,82,0.1)",
  red:         "#C0392B",
  redLight:    "rgba(192,57,43,0.1)",
  orange:      "#D4770A",
  orangeLight: "rgba(212,119,10,0.1)",
  purple:      "#6B46C1",

  // 그림자
  shadow:   "0 1px 4px rgba(0,0,0,0.06), 0 4px 16px rgba(0,0,0,0.05)",
  shadowMd: "0 4px 24px rgba(0,0,0,0.09), 0 1px 6px rgba(0,0,0,0.05)",
  shadowLg: "0 12px 48px rgba(0,0,0,0.13), 0 3px 12px rgba(0,0,0,0.07)",

  // 형태
  radius:   "10px",
  radiusSm: "7px",
  radiusXl: "16px",

  // 폰트 — 전문적 산세리프
  font: "'Pretendard', 'Apple SD Gothic Neo', -apple-system, BlinkMacSystemFont, 'Malgun Gothic', sans-serif",
};

// Card 기본 스타일
const CARD_STYLE = {
  background:  "#FFFFFF",
  borderRadius: T.radius,
  border:      `1px solid rgba(0,0,0,0.09)`,
  boxShadow:   T.shadow,
};

/* ─────────────────────────────────────────
   반응형 훅
───────────────────────────────────────── */
const useIsMobile = () => {
  const [mobile, setMobile] = useState(window.innerWidth < 768);
  useEffect(() => {
    const fn = () => setMobile(window.innerWidth < 768);
    window.addEventListener("resize", fn);
    return () => window.removeEventListener("resize", fn);
  }, []);
  return mobile;
};

/* ─────────────────────────────────────────
   SVG 아이콘 컴포넌트 (이모티콘 대체)
   Stroke 기반 — 전문적이고 일관된 스타일
───────────────────────────────────────── */
const Icon = ({ name, size=16, color="currentColor", strokeWidth=1.6 }) => {
  const s = { width:size, height:size, display:"inline-block", flexShrink:0 };
  const p = { fill:"none", stroke:color, strokeWidth, strokeLinecap:"round", strokeLinejoin:"round" };
  const paths = {
    // 메뉴 아이콘
    balanceSheet: <svg style={s} viewBox="0 0 24 24"><rect {...p} x="3" y="3" width="18" height="18" rx="2"/><line {...p} x1="3" y1="9" x2="21" y2="9"/><line {...p} x1="9" y1="9" x2="9" y2="21"/></svg>,
    income:       <svg style={s} viewBox="0 0 24 24"><polyline {...p} points="22 7 13.5 15.5 8.5 10.5 2 17"/><polyline {...p} points="16 7 22 7 22 13"/></svg>,
    vat:          <svg style={s} viewBox="0 0 24 24"><path {...p} d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline {...p} points="14 2 14 8 20 8"/><line {...p} x1="16" y1="13" x2="8" y2="13"/><line {...p} x1="16" y1="17" x2="8" y2="17"/><polyline {...p} points="10 9 9 9 8 9"/></svg>,
    tax:          <svg style={s} viewBox="0 0 24 24"><path {...p} d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline {...p} points="3.27 6.96 12 12.01 20.73 6.96"/><line {...p} x1="12" y1="22.08" x2="12" y2="12"/></svg>,
    calendar:     <svg style={s} viewBox="0 0 24 24"><rect {...p} x="3" y="4" width="18" height="18" rx="2" ry="2"/><line {...p} x1="16" y1="2" x2="16" y2="6"/><line {...p} x1="8" y1="2" x2="8" y2="6"/><line {...p} x1="3" y1="10" x2="21" y2="10"/></svg>,
    settings:     <svg style={s} viewBox="0 0 24 24"><circle {...p} cx="12" cy="12" r="3"/><path {...p} d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>,
    // UI 아이콘
    menu:         <svg style={s} viewBox="0 0 24 24"><line {...p} x1="3" y1="6" x2="21" y2="6"/><line {...p} x1="3" y1="12" x2="21" y2="12"/><line {...p} x1="3" y1="18" x2="21" y2="18"/></svg>,
    logout:       <svg style={s} viewBox="0 0 24 24"><path {...p} d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline {...p} points="16 17 21 12 16 7"/><line {...p} x1="21" y1="12" x2="9" y2="12"/></svg>,
    logo:         <svg style={s} viewBox="0 0 24 24"><line {...p} x1="18" y1="20" x2="18" y2="10"/><line {...p} x1="12" y1="20" x2="12" y2="4"/><line {...p} x1="6" y1="20" x2="6" y2="14"/></svg>,
    check:        <svg style={s} viewBox="0 0 24 24"><polyline {...p} points="20 6 9 17 4 12"/></svg>,
    clock:        <svg style={s} viewBox="0 0 24 24"><circle {...p} cx="12" cy="12" r="10"/><polyline {...p} points="12 6 12 12 16 14"/></svg>,
    inbox:        <svg style={s} viewBox="0 0 24 24"><polyline {...p} points="22 12 16 12 14 15 10 15 8 12 2 12"/><path {...p} d="M5.45 5.11L2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/></svg>,
    upload:       <svg style={s} viewBox="0 0 24 24"><polyline {...p} points="16 16 12 12 8 16"/><line {...p} x1="12" y1="12" x2="12" y2="21"/><path {...p} d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/></svg>,
    building:     <svg style={s} viewBox="0 0 24 24"><rect {...p} x="4" y="2" width="16" height="20" rx="0"/><line {...p} x1="9" y1="22" x2="9" y2="2"/><line {...p} x1="4" y1="7" x2="9" y2="7"/><line {...p} x1="4" y1="12" x2="9" y2="12"/><line {...p} x1="4" y1="17" x2="9" y2="17"/><line {...p} x1="14" y1="7" x2="20" y2="7"/><line {...p} x1="14" y1="12" x2="20" y2="12"/><line {...p} x1="14" y1="17" x2="20" y2="17"/></svg>,
    users:        <svg style={s} viewBox="0 0 24 24"><path {...p} d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle {...p} cx="9" cy="7" r="4"/><path {...p} d="M23 21v-2a4 4 0 0 0-3-3.87"/><path {...p} d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>,
    alert:        <svg style={s} viewBox="0 0 24 24"><circle {...p} cx="12" cy="12" r="10"/><line {...p} x1="12" y1="8" x2="12" y2="12"/><line {...p} x1="12" y1="16" x2="12.01" y2="16"/></svg>,
    download:     <svg style={s} viewBox="0 0 24 24"><path {...p} d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline {...p} points="7 10 12 15 17 10"/><line {...p} x1="12" y1="15" x2="12" y2="3"/></svg>,
    list:         <svg style={s} viewBox="0 0 24 24"><line {...p} x1="8" y1="6" x2="21" y2="6"/><line {...p} x1="8" y1="12" x2="21" y2="12"/><line {...p} x1="8" y1="18" x2="21" y2="18"/><line {...p} x1="3" y1="6" x2="3.01" y2="6"/><line {...p} x1="3" y1="12" x2="3.01" y2="12"/><line {...p} x1="3" y1="18" x2="3.01" y2="18"/></svg>,
    timeline:     <svg style={s} viewBox="0 0 24 24"><line {...p} x1="12" y1="2" x2="12" y2="22"/><circle {...p} cx="12" cy="6" r="2" fill={color}/><circle {...p} cx="12" cy="12" r="2" fill={color}/><circle {...p} cx="12" cy="18" r="2" fill={color}/><line {...p} x1="12" y1="6" x2="20" y2="6"/><line {...p} x1="12" y1="12" x2="20" y2="12"/><line {...p} x1="12" y1="18" x2="20" y2="18"/></svg>,
    grid:         <svg style={s} viewBox="0 0 24 24"><rect {...p} x="3" y="3" width="7" height="7"/><rect {...p} x="14" y="3" width="7" height="7"/><rect {...p} x="3" y="14" width="7" height="7"/><rect {...p} x="14" y="14" width="7" height="7"/></svg>,
    dot:          <svg style={s} viewBox="0 0 24 24"><circle fill={color} cx="12" cy="12" r="4"/></svg>,
  };
  return paths[name] || <svg style={s} viewBox="0 0 24 24"><circle {...p} cx="12" cy="12" r="10"/></svg>;
};

/* ─────────────────────────────────────────
   로컬 스토리지 스토어
───────────────────────────────────────── */
const KEY = "finportal_v2";
const DEFAULT = {
  users: [
    { id:1, name:"홍길동", email:"hong@example.com", password:"1234", role:"customer", status:"approved", businesses:["123-45-67890","987-65-43210"], phone:"010-1234-5678", registeredAt:"2025-01-10T09:00:00Z", memo:"" },
    { id:2, name:"김철수", email:"kim@example.com", password:"1234", role:"customer", status:"pending", businesses:["111-22-33333"], phone:"010-9876-5432", registeredAt:"2025-05-20T14:30:00Z", memo:"빠른 승인 부탁드립니다" },
    { id:99, name:"관리자", email:"admin@cpa.com", password:"admin1234", role:"admin", status:"approved", businesses:[] },
  ],
  businesses: {
    "123-45-67890": { name:"(주)한국상사", type:"법인", representative:"홍길동" },
    "987-65-43210": { name:"홍길동 부동산", type:"개인", representative:"홍길동" },
    "111-22-33333": { name:"철수네 카페", type:"개인", representative:"김철수" },
  },
  financial: {
    "123-45-67890": {
      "2024": {
        "재무상태표": { 기간:"2024년 12월 31일", rows:[
          {계정과목:"유동자산",당기:288000000,전기:228000000,level:0,type:"header"},
          {계정과목:"현금및현금성자산",당기:85000000,전기:62000000,level:1,type:"item"},
          {계정과목:"매출채권",당기:120000000,전기:98000000,level:1,type:"item"},
          {계정과목:"재고자산",당기:65000000,전기:54000000,level:1,type:"item"},
          {계정과목:"기타유동자산",당기:18000000,전기:14000000,level:1,type:"item"},
          {계정과목:"비유동자산",당기:415000000,전기:385000000,level:0,type:"header"},
          {계정과목:"유형자산",당기:350000000,전기:320000000,level:1,type:"item"},
          {계정과목:"무형자산",당기:25000000,전기:30000000,level:1,type:"item"},
          {계정과목:"기타비유동자산",당기:40000000,전기:35000000,level:1,type:"item"},
          {계정과목:"자산총계",당기:703000000,전기:613000000,level:0,type:"total"},
          {계정과목:"유동부채",당기:197000000,전기:186000000,level:0,type:"header"},
          {계정과목:"매입채무",당기:95000000,전기:78000000,level:1,type:"item"},
          {계정과목:"단기차입금",당기:80000000,전기:90000000,level:1,type:"item"},
          {계정과목:"기타유동부채",당기:22000000,전기:18000000,level:1,type:"item"},
          {계정과목:"비유동부채",당기:180000000,전기:205000000,level:0,type:"header"},
          {계정과목:"장기차입금",당기:150000000,전기:180000000,level:1,type:"item"},
          {계정과목:"기타비유동부채",당기:30000000,전기:25000000,level:1,type:"item"},
          {계정과목:"부채총계",당기:377000000,전기:391000000,level:0,type:"total"},
          {계정과목:"자본금",당기:100000000,전기:100000000,level:1,type:"item"},
          {계정과목:"이익잉여금",당기:226000000,전기:122000000,level:1,type:"item"},
          {계정과목:"자본총계",당기:326000000,전기:222000000,level:0,type:"total"},
        ]},
        "손익계산서": { 기간:"2024.01.01 ~ 12.31", rows:[
          {계정과목:"매출액",당기:980000000,전기:820000000,level:0,type:"header"},
          {계정과목:"매출원가",당기:620000000,전기:540000000,level:1,type:"item"},
          {계정과목:"매출총이익",당기:360000000,전기:280000000,level:0,type:"subtotal"},
          {계정과목:"판매비및관리비",당기:180000000,전기:165000000,level:0,type:"header"},
          {계정과목:"급여",당기:95000000,전기:85000000,level:1,type:"item"},
          {계정과목:"임차료",당기:24000000,전기:24000000,level:1,type:"item"},
          {계정과목:"감가상각비",당기:18000000,전기:20000000,level:1,type:"item"},
          {계정과목:"광고선전비",당기:12000000,전기:8000000,level:1,type:"item"},
          {계정과목:"기타",당기:31000000,전기:28000000,level:1,type:"item"},
          {계정과목:"영업이익",당기:180000000,전기:115000000,level:0,type:"subtotal"},
          {계정과목:"영업외수익",당기:8000000,전기:5000000,level:1,type:"item"},
          {계정과목:"영업외비용",당기:15000000,전기:18000000,level:1,type:"item"},
          {계정과목:"법인세차감전이익",당기:173000000,전기:102000000,level:0,type:"subtotal"},
          {계정과목:"법인세비용",당기:30000000,전기:17000000,level:1,type:"item"},
          {계정과목:"당기순이익",당기:143000000,전기:85000000,level:0,type:"total"},
        ]},
        "부가세신고": { rows:[
          {기수:"1기 예정",기간:"2024.01~03",신고기한:"2024-04-25",매출세액:18500000,매입세액:12300000,납부세액:6200000,상태:"완료"},
          {기수:"1기 확정",기간:"2024.04~06",신고기한:"2024-07-25",매출세액:22100000,매입세액:15800000,납부세액:6300000,상태:"완료"},
          {기수:"2기 예정",기간:"2024.07~09",신고기한:"2024-10-25",매출세액:24500000,매입세액:16200000,납부세액:8300000,상태:"완료"},
          {기수:"2기 확정",기간:"2024.10~12",신고기한:"2025-01-25",매출세액:27800000,매입세액:18900000,납부세액:8900000,상태:"완료"},
        ]},
        "법인세신고": { 사업연도:"2024.01.01 ~ 12.31", 신고기한:"2025-03-31", 과세표준:173000000, 산출세액:30000000, 공제감면세액:2500000, 납부세액:27500000, 신고상태:"완료" },
      }
    }
  }
};
/* ─────────────────────────────────────────
   Supabase 연동
───────────────────────────────────────── */
const SUPABASE_URL = "https://mjnkocgoevuntspepwrf.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1qbmtvY2dvZXZ1bnRzcGVwd3JmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk5NjE1ODEsImV4cCI6MjA5NTUzNzU4MX0.h9yQEO0CmJOhYMHWz3oxAbuxWCMVqDIaj_07l1MvGNM";

const supa = async (path, options={}) => {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    headers: {
      "apikey":        SUPABASE_KEY,
      "Authorization": `Bearer ${SUPABASE_KEY}`,
      "Content-Type":  "application/json",
      "Prefer":        options.prefer || "return=representation",
      ...options.headers,
    },
    ...options,
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(err);
  }
  const text = await res.text();
  return text ? JSON.parse(text) : null;
};

// GET  /table?filter
const supaGet  = (table, query="") => supa(`${table}?${query}`);
// POST /table  body
const supaPost = (table, body) => supa(table, { method:"POST", body:JSON.stringify(body) });
// PATCH /table?filter body
const supaPatch = (table, query, body) =>
  supa(`${table}?${query}`, { method:"PATCH", body:JSON.stringify(body), prefer:"return=representation" });
// UPSERT
const supaUpsert = (table, body) =>
  supa(table, { method:"POST", body:JSON.stringify(body), headers:{ "Prefer":"resolution=merge-duplicates,return=representation", "Content-Type":"application/json" } });

const db = {
  // ── 이메일 존재 확인 (비밀번호 찾기용)
  checkEmail: async (email) => {
    const rows = await supaGet("users", `email=eq.${encodeURIComponent(email)}&select=id`);
    return !!(rows?.length);
  },

  // ── 로그인
  login: async (email, password) => {
    // 어드민 계정 하드코딩 폴백 (Supabase DB에 없는 경우 대비)
    if (email === "admin@cpa.com" && password === "admin1234") {
      return { ok:true, user:{ id:99, name:"관리자", email:"admin@cpa.com", role:"admin", status:"approved", phone:"", businesses:[] } };
    }
    const rows = await supaGet("users", `email=eq.${encodeURIComponent(email)}&password=eq.${encodeURIComponent(password)}&select=*`);
    if (!rows?.length) return { ok:false, msg:"이메일 또는 비밀번호가 올바르지 않습니다." };
    const u = rows[0];
    if (u.status === "pending")  return { ok:false, msg:"관리자 승인 대기 중입니다." };
    if (u.status === "rejected") return { ok:false, msg:"가입이 거절되었습니다." };
    // 사업자 목록 조회
    const ubRows = await supaGet("user_businesses", `user_id=eq.${u.id}&select=biz_no`);
    const businesses = (ubRows||[]).map(r=>r.biz_no);
    return { ok:true, user:{ id:u.id, name:u.name, email:u.email, role:u.role, status:u.status, phone:u.phone, businesses } };
  },

  // ── 회원가입
  register: async (data) => {
    const exist = await supaGet("users", `email=eq.${encodeURIComponent(data.email)}&select=id`);
    if (exist?.length) return { ok:false, msg:"이미 사용 중인 이메일입니다." };
    const newUser = await supaPost("users", {
      name:data.name, email:data.email, password:data.password,
      phone:data.phone||"", memo:data.memo||"",
      role:"customer", status:"pending",
    });
    const userId = newUser?.[0]?.id;
    if (userId && data.businesses?.length) {
      for (const biz_no of data.businesses) {
        const trimmed = biz_no.trim();
        if (!trimmed) continue;
        // user_businesses 테이블에 연결
        await supaPost("user_businesses", { user_id:userId, biz_no:trimmed }).catch(()=>{});
        // businesses 테이블에도 등록 (상호명 미입력 시 사업자번호를 임시 상호로)
        const bizName = (data.bizNames && data.bizNames[trimmed]) || trimmed;
        await supaUpsert("businesses", { biz_no:trimmed, name:bizName, type:"개인", representative:data.name||"" }).catch(()=>{});
      }
    }
    return { ok:true };
  },

  // ── 전체 사업자번호 목록 (user_businesses 포함 — 업로드 패널용)
  getAllBizNos: async () => {
    // businesses 테이블 + user_businesses에 있지만 businesses에 없는 것도 포함
    const bizRows = await supaGet("businesses", "select=*&order=name");
    const ubRows  = await supaGet("user_businesses", "select=biz_no");
    const bizMap = {};
    (bizRows||[]).forEach(r=>{ bizMap[r.biz_no]={ name:r.name, type:r.type, representative:r.representative }; });
    // user_businesses에만 있는 사업자번호 추가 (상호 미등록 상태)
    (ubRows||[]).forEach(r=>{ if(!bizMap[r.biz_no]) bizMap[r.biz_no]={ name:r.biz_no, type:"-", representative:"" }; });
    return bizMap;
  },

  // ── 사용자 목록 (관리자용)
  getUsers: async () => {
    const users = await supaGet("users", "role=eq.customer&select=*&order=registered_at.desc");
    const ubs   = await supaGet("user_businesses", "select=user_id,biz_no");
    return (users||[]).map(u=>({
      ...u,
      registeredAt: u.registered_at,
      businesses: (ubs||[]).filter(r=>r.user_id===u.id).map(r=>r.biz_no),
    }));
  },

  // ── 회원 정보 수정 (어드민)
  updateUser: async (id, data) => {
    const body = {};
    if(data.name!==undefined) body.name=data.name;
    if(data.email!==undefined) body.email=data.email;
    if(data.phone!==undefined) body.phone=data.phone;
    if(data.password!==undefined && data.password) body.password=data.password;
    if(data.memo!==undefined) body.memo=data.memo;
    if(data.status!==undefined) body.status=data.status;
    await supaPatch("users", `id=eq.${id}`, body);
  },

  // ── 회원 삭제
  deleteUser: async (id) => {
    // user_businesses 연결 먼저 삭제
    await supa(`user_businesses?user_id=eq.${id}`, { method:"DELETE", prefer:"" });
    await supa(`users?id=eq.${id}`, { method:"DELETE", prefer:"" });
  },

  // ── 회원 상태 변경
  setStatus: async (id, status) => {
    await supaPatch("users", `id=eq.${id}`, { status });
  },

  // ── 회원 사업자 목록 갱신
  setBizList: async (userId, list) => {
    // 기존 삭제 후 재삽입
    await supa(`user_businesses?user_id=eq.${userId}`, { method:"DELETE", prefer:"" });
    for (const biz_no of list) {
      await supaPost("user_businesses", { user_id:userId, biz_no }).catch(()=>{});
    }
  },

  // ── 사업자 조회
  getBiz: async (no) => {
    const rows = await supaGet("businesses", `biz_no=eq.${encodeURIComponent(no)}&select=*`);
    return rows?.[0] || null;
  },

  // ── 전체 사업자 목록
  allBiz: async () => {
    const rows = await supaGet("businesses", "select=*&order=name");
    const obj = {};
    (rows||[]).forEach(r=>{ obj[r.biz_no]={ name:r.name, type:r.type, representative:r.representative }; });
    return obj;
  },

  // ── 사업자 추가
  addBiz: async (no, info) => {
    try {
      await supaUpsert("businesses", { biz_no:no, name:info.name, type:info.type, representative:info.representative||"" });
    } catch(e) {
      // 중복 키 오류(409) 등은 무시하고 PATCH로 재시도
      await supaPatch("businesses", `biz_no=eq.${encodeURIComponent(no)}`, { name:info.name, type:info.type, representative:info.representative||"" });
    }
  },

  // ── 재무 데이터 조회
  getDoc: async (no, yr, type) => {
    const rows = await supaGet("financial_docs",
      `biz_no=eq.${encodeURIComponent(no)}&year=eq.${yr}&doc_type=eq.${encodeURIComponent(type)}&select=data`);
    return rows?.[0]?.data || null;
  },

  // ── 사업자의 데이터 있는 연도 목록
  getYears: async (no) => {
    const rows = await supaGet("financial_docs",
      `biz_no=eq.${encodeURIComponent(no)}&select=year&order=year.desc`);
    return [...new Set((rows||[]).map(r=>r.year))];
  },

  // ── 재무 데이터 저장 (upsert)
  saveDoc: async (no, yr, type, data) => {
    await supaUpsert("financial_docs", { biz_no:no, year:yr, doc_type:type, data });
  },
};

/* ─────────────────────────────────────────
   숫자 포맷 (#,##0 ; (#,##0) ; -)
───────────────────────────────────────── */
const F = (n) => {
  if(n===null||n===undefined) return "-";
  if(n===0) return "-";
  const a=Math.abs(n), f=a.toLocaleString("ko-KR");
  return n<0?`(${f})`:f;
};
const FW = (n) => n===0||!n ? "-" : F(n)+"원";

// 천원 단위 절사 (모바일용) — 1,000원 미만 버림
// 예: 85,000,000 → 85,000  /  -5,000,000 → (5,000)
const FK = (n) => {
  if(n===null||n===undefined||n===0) return "-";
  const thou = Math.round(n / 1000);
  if(thou===0) return "-";
  const a = Math.abs(thou);
  const f = a.toLocaleString("ko-KR");
  return thou < 0 ? `(${f})` : f;
};

/* ─────────────────────────────────────────
   위하고 Excel 파서 — 행별 동적 컬럼 선택
   위하고T 실제 구조:
     합계/소계행 → col2(당기합계), col4(전기합계)
     세부항목행  → col1(당기세부), col3(전기세부)
     감가상각누계행 → col2(당기순액), col4(전기순액) + 직전 자산행 순액 보정
───────────────────────────────────────── */
const parseExcel = (buffer, docType) => {
  const wb = XLSX.read(buffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:"", raw:false });

  const toNum = (v) => {
    if(v===null||v===undefined||v==="") return 0;
    if(typeof v === "number") return v;
    const s = String(v).replace(/,/g,"").replace(/\s/g,"").replace(/원/g,"");
    const neg = s.match(/^\((.+)\)$/);
    const n = parseFloat(neg ? neg[1] : s);
    return isNaN(n) ? 0 : (neg ? -n : n);
  };
  const hasV = (v) => v !== null && v !== undefined && String(v).trim() !== "" && String(v).trim() !== "0";
  if (docType === "재무상태표" || docType === "손익계산서") {
    let dataStart = 0;
    for (let i = 0; i < Math.min(raw.length, 15); i++) {
      const row = raw[i].map(c => String(c).trim());
      if (row.some(c => /^(과\s*목|계정과목|항\s*목)$/.test(c))) {
        dataStart = i + 1; break;
      }
    }
    // "금액" 서브헤더 행 스킵
    if (dataStart < raw.length) {
      const sub = raw[dataStart].map(c => String(c).trim());
      if (sub.every(c => c === "" || /^금\s*액$/.test(c))) dataStart++;
    }

    // 기간 텍스트
    let 기간 = "";
    for (let i = 0; i < Math.min(dataStart, raw.length); i++) {
      const line = raw[i].join(" ");
      if (/20\d{2}/.test(line)) { 기간 = line.replace(/\s+/g," ").trim(); break; }
    }

    const totalKw    = /총계|당기순이익|순이익|자본총|부채총|자산총|부채및자본총계/;
    const subtotalKw = /총이익|영업이익|차감전|소\s*계|매출총이익/;
    const skipKw     = /^(자산|부채|자본)$/;
    const accumKw    = /감가상각누계액|대손충당금/;

    const rows = [];
    let pendingIdx = -1; // 취득액만 있고 순액 미확정인 유형자산 행 인덱스

    for (let i = dataStart; i < raw.length; i++) {
      const row = raw[i];
      if (!row || row.every(v => v === "" || v === null)) { pendingIdx = -1; continue; }

      const acRaw = String(row[0] || "").trim();
      if (!acRaw || acRaw === "0") continue;
      if (/^(당기|전기)\s*:/.test(acRaw)) continue;
      if (/^\(당기순이익\)$/.test(acRaw)) { pendingIdx = -1; continue; }
      if (skipKw.test(acRaw)) { pendingIdx = -1; continue; }

      // 감가상각누계액 행 → 직전 pendingIdx 자산행 순액 보정 후 스킵
      if (accumKw.test(acRaw)) {
        if (pendingIdx >= 0) {
          const netCur  = toNum(row[2]);
          const netPrev = toNum(row[4]);
          if (netCur !== 0 || netPrev !== 0) {
            rows[pendingIdx].당기 = netCur;
            rows[pendingIdx].전기 = netPrev;
          }
          pendingIdx = -1;
        }
        continue;
      }

      // 행별 동적 컬럼 선택
      const v1 = row[1], v2 = row[2], v3 = row[3], v4 = row[4];
      let cur = 0, prev = 0;
      let isPending = false;

      if (hasV(v2) || hasV(v4)) {
        // 합계/그룹/일반 세부 행 (col2·col4에 값)
        cur  = toNum(v2);
        prev = toNum(v4);
        pendingIdx = -1;
      } else if (hasV(v1) || hasV(v3)) {
        // 세부항목 행 (col1·col3에만 값) — 손익계산서 세부 또는 유형자산 취득액
        cur  = toNum(v1);
        prev = toNum(v3);
        isPending = true; // 재무상태표에서 다음 누계액 행으로 순액 보정 대기
      } else {
        continue;
      }

      const rawAc = String(row[0] || "");
      const lead  = (rawAc.match(/^(\s+)/) || ["",""])[1].length;
      const level = lead > 10 ? 2 : lead > 3 ? 1 : 0;

      let type = "item";
      if (totalKw.test(acRaw))         type = "total";
      else if (subtotalKw.test(acRaw)) type = "subtotal";
      else if (level === 0)            type = "header";

      if (isPending) pendingIdx = rows.length;
      else pendingIdx = -1;

      rows.push({ 계정과목: acRaw, 당기: cur, 전기: prev, level, type });
    }

    return { 기간, rows };
  }

  if (docType === "부가세신고") {
    // 부가세: 기수/신고기한/매출세액/매입세액/납부세액 패턴
    let hIdx=-1, cols={};
    for (let i=0; i<Math.min(raw.length,15); i++) {
      const row = raw[i].map(c=>String(c).trim());
      if (row.some(c=>/매출세액|매출\s*세/.test(c))) {
        hIdx=i;
        cols = {
          기수: row.findIndex(c=>/기수|구분|기간/.test(c)),
          신고기한: row.findIndex(c=>/기한|신고일/.test(c)),
          기간: row.findIndex(c=>/^기간$|대상기간/.test(c)),
          매출세액: row.findIndex(c=>/매출세액|매출\s*세/.test(c)),
          매입세액: row.findIndex(c=>/매입세액|매입\s*세/.test(c)),
          납부세액: row.findIndex(c=>/납부세액|납부\s*세|차감납부/.test(c)),
        };
        break;
      }
    }
    const rows = [];
    for (let i=(hIdx>=0?hIdx+1:1); i<raw.length; i++) {
      const r = raw[i];
      const 기수 = String(r[cols.기수>=0?cols.기수:0]||"").trim();
      if (!기수 || 기수==="합계") continue;
      rows.push({
        기수,
        기간: cols.기간>=0?String(r[cols.기간]||"").trim():"",
        신고기한: cols.신고기한>=0?String(r[cols.신고기한]||"").trim():"",
        매출세액: toNum(r[cols.매출세액>=0?cols.매출세액:1]),
        매입세액: toNum(r[cols.매입세액>=0?cols.매입세액:2]),
        납부세액: toNum(r[cols.납부세액>=0?cols.납부세액:3]),
        상태: "완료",
      });
    }
    return { rows };
  }

  if (docType === "법인세신고") {
    // 법인세: key-value 형태의 시트 파싱
    const result = { 사업연도:"", 신고기한:"", 과세표준:0, 산출세액:0, 공제감면세액:0, 납부세액:0, 신고상태:"완료" };
    for (const row of raw) {
      const line = row.map(c=>String(c).trim());
      for (let i=0; i<line.length-1; i++) {
        const k=line[i], v=line[i+1];
        if (/사업연도|회계기간/.test(k)) result.사업연도=v;
        else if (/신고기한/.test(k)) result.신고기한=v;
        else if (/과세표준/.test(k)) result.과세표준=toNum(v);
        else if (/산출세액/.test(k)) result.산출세액=toNum(v);
        else if (/공제|감면/.test(k)) result.공제감면세액=toNum(v);
        else if (/납부세액|최종납부/.test(k)) result.납부세액=toNum(v);
      }
    }
    return result;
  }

  return null;
};

/* ─────────────────────────────────────────
   공통 컴포넌트
───────────────────────────────────────── */
const Card = ({ children, style={}, onClick }) => (
  <div onClick={onClick} style={{
    background:T.surface, borderRadius:T.radius,
    border:`1px solid ${T.border}`, boxShadow:T.shadow,
    backdropFilter:"blur(20px)", WebkitBackdropFilter:"blur(20px)",
    transition:"all 0.2s ease", ...style,
    ...(onClick?{cursor:"pointer"}:{}),
  }}
    onMouseEnter={e=>{if(onClick){e.currentTarget.style.boxShadow=T.shadowMd;e.currentTarget.style.transform="translateY(-1px)";}}}
    onMouseLeave={e=>{if(onClick){e.currentTarget.style.boxShadow=T.shadow;e.currentTarget.style.transform="translateY(0)";}}}
  >{children}</div>
);

const Pill = ({ label, color=T.blue, bg=T.blueLight }) => (
  <span style={{display:"inline-block",background:bg,color,fontSize:"11px",fontWeight:"600",padding:"3px 10px",borderRadius:"20px",letterSpacing:"-0.1px"}}>{label}</span>
);

const KPI = ({ label, value, sub, color=T.blue, icon }) => (
  <Card style={{padding:"20px 22px"}}>
    <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:"12px"}}>
      <span style={{fontSize:"12px",fontWeight:"500",color:T.textSub,letterSpacing:"-0.1px"}}>{label}</span>
      {icon&&<span style={{fontSize:"18px"}}>{icon}</span>}
    </div>
    <div style={{fontSize:"22px",fontWeight:"700",color:T.text,letterSpacing:"-0.8px",fontVariantNumeric:"tabular-nums"}}>{value}</div>
    {sub&&<div style={{fontSize:"12px",color,fontWeight:"500",marginTop:"4px"}}>{sub}</div>}
  </Card>
);

const Btn = ({ children, onClick, variant="primary", disabled=false, size="md", style={} }) => {
  const base = { border:"none", cursor:disabled?"not-allowed":"pointer", fontFamily:T.font, fontWeight:"600", letterSpacing:"-0.2px", transition:"all 0.15s ease", opacity:disabled?0.5:1, ...style };
  const sizes = { sm:{padding:"6px 14px",fontSize:"13px",borderRadius:"8px"}, md:{padding:"10px 20px",fontSize:"14px",borderRadius:T.radiusSm}, lg:{padding:"14px 28px",fontSize:"15px",borderRadius:T.radius} };
  const variants = {
    primary:{background:T.blue,color:"#fff",boxShadow:"0 1px 3px rgba(0,113,227,0.3)"},
    secondary:{background:"rgba(0,0,0,0.06)",color:T.text},
    ghost:{background:"transparent",color:T.blue},
    danger:{background:T.red,color:"#fff"},
  };
  return <button onClick={disabled?undefined:onClick} style={{...base,...sizes[size],...variants[variant]}}>{children}</button>;
};

const Input = ({ value, onChange, placeholder, type="text", style={}, onKeyDown }) => (
  <input type={type} value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder}
    onKeyDown={onKeyDown}
    style={{width:"100%",padding:"10px 14px",borderRadius:T.radiusSm,border:`1.5px solid ${T.border}`,background:"rgba(255,255,255,0.9)",color:T.text,fontSize:"14px",fontFamily:T.font,outline:"none",boxSizing:"border-box",transition:"border-color 0.15s",...style}}
    onFocus={e=>e.target.style.borderColor=T.blue}
    onBlur={e=>e.target.style.borderColor=T.border}
  />
);

const Label = ({ children }) => (
  <div style={{fontSize:"12px",fontWeight:"600",color:T.textSub,marginBottom:"6px",letterSpacing:"-0.1px"}}>{children}</div>
);

/* ─────────────────────────────────────────
   로그인
───────────────────────────────────────── */
function ForgotPassword({ onBack }) {
  const [step,setStep]=useState(1); // 1=이메일입력, 2=완료
  const [email,setEmail]=useState(""), [loading,setLoading]=useState(false), [err,setErr]=useState("");
  const submit=()=>{
    setErr("");
    if(!email){setErr("이메일을 입력해주세요.");return;}
    setLoading(true);
    // 이메일 존재 여부 확인
    db.checkEmail(email)
      .then(exists=>{ setLoading(false); if(exists) setStep(2); else setErr("등록되지 않은 이메일입니다."); })
      .catch(()=>{ setLoading(false); setErr("서버 오류가 발생했습니다."); });
  };
  return (
    <div style={{minHeight:"100vh",fontFamily:T.font,display:"flex",alignItems:"center",justifyContent:"center",
      background:`linear-gradient(160deg,${T.bgDeep} 0%,${T.bgDeepAlt} 60%,#0D1829 100%)`}}>
      <div style={{width:"100%",maxWidth:"380px",padding:"24px 20px"}}>
        <div style={{background:"rgba(255,255,255,0.97)",borderRadius:T.radiusXl,padding:"40px 36px",boxShadow:"0 24px 64px rgba(0,0,0,0.3)"}}>
          {step===1?(
            <>
              <h3 style={{color:T.navy,fontSize:"18px",fontWeight:"800",margin:"0 0 4px"}}>비밀번호 찾기</h3>
              <p style={{color:T.textMuted,fontSize:"12px",margin:"0 0 22px"}}>가입하신 이메일을 입력하면 담당자가 안내해드립니다.</p>
              <div style={{marginBottom:"16px"}}><Label>이메일</Label><Input value={email} onChange={setEmail} placeholder="example@email.com" type="email" onKeyDown={ev=>ev.key==="Enter"&&submit()}/></div>
              {err&&<div style={{background:T.redLight,border:`1px solid rgba(192,57,43,0.2)`,borderRadius:T.radiusSm,padding:"10px 13px",marginBottom:"16px",color:T.red,fontSize:"13px",fontWeight:"500"}}>{err}</div>}
              <button onClick={submit} disabled={loading} style={{width:"100%",padding:"12px",borderRadius:T.radiusSm,border:"none",background:`linear-gradient(135deg,${T.bgDeep},${T.bgDeepAlt})`,color:"#fff",fontSize:"14px",fontWeight:"700",cursor:"pointer",fontFamily:T.font,marginBottom:"14px",opacity:loading?0.7:1}}>{loading?"확인 중…":"확인"}</button>
              <div style={{textAlign:"center"}}><button onClick={onBack} style={{background:"none",border:"none",color:T.textSub,fontSize:"13px",cursor:"pointer"}}>← 로그인으로 돌아가기</button></div>
            </>
          ):(
            <>
              <div style={{textAlign:"center",marginBottom:"20px"}}>
                <div style={{width:"56px",height:"56px",borderRadius:"50%",background:T.greenLight,display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 16px"}}>
                  <Icon name="check" size={28} color={T.green} strokeWidth={2}/>
                </div>
                <h3 style={{color:T.navy,fontSize:"18px",fontWeight:"800",margin:"0 0 8px"}}>접수 완료</h3>
                <p style={{color:T.textSub,fontSize:"13px",lineHeight:"1.7",margin:"0 0 24px"}}>담당 회계사에게 문의 내역이 전달되었습니다.<br/>카카오톡 채널 또는 이메일로 안내해드리겠습니다.</p>
                <Btn onClick={onBack} size="md">로그인으로 돌아가기</Btn>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

function Login({ onLogin, onGo }) {
  const SAVED_EMAIL_KEY = "finportal_saved_email";
  const [e,setE]=useState(()=>localStorage.getItem(SAVED_EMAIL_KEY)||"");
  const [p,setP]=useState(""), [err,setErr]=useState(""), [loading,setLoading]=useState(false);
  const [saveEmail,setSaveEmail]=useState(()=>!!localStorage.getItem(SAVED_EMAIL_KEY));
  const [showForgot,setShowForgot]=useState(false);
  const go = () => {
    setErr(""); if(!e||!p){setErr("이메일과 비밀번호를 입력해주세요.");return;}
    if(saveEmail) localStorage.setItem(SAVED_EMAIL_KEY,e);
    else localStorage.removeItem(SAVED_EMAIL_KEY);
    setLoading(true);
    db.login(e,p)
      .then(r=>{ setLoading(false); if(r.ok) onLogin(r.user); else setErr(r.msg); })
      .catch(()=>{ setLoading(false); setErr("서버 연결 오류가 발생했습니다."); });
  };
  const isMobile = useIsMobile();
  if(showForgot) return <ForgotPassword onBack={()=>setShowForgot(false)}/>;
  return (
    <div style={{
      minHeight:"100vh",fontFamily:T.font,display:"flex",
      background:`linear-gradient(160deg, ${T.bgDeep} 0%, ${T.bgDeepAlt} 60%, #0D1829 100%)`,
    }}>
      {/* 배경 장식 */}
      <div style={{position:"fixed",inset:0,overflow:"hidden",pointerEvents:"none"}}>
        <div style={{position:"absolute",top:"-10%",right:"-5%",width:"500px",height:"500px",borderRadius:"50%",background:"radial-gradient(circle,rgba(201,168,76,0.07) 0%,transparent 65%)"}}/>
        <div style={{position:"absolute",bottom:"-15%",left:"-8%",width:"600px",height:"600px",borderRadius:"50%",background:"radial-gradient(circle,rgba(30,95,173,0.12) 0%,transparent 65%)"}}/>

      </div>



      {/* 로그인 폼 — 모바일은 전체, PC는 우측 */}
      <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",padding:isMobile?"24px 20px":"40px 24px"}}>
        <div style={{width:"100%",maxWidth:"380px"}}>
          {/* 로고 (항상 표시) */}
          <div style={{textAlign:"center",marginBottom:"28px"}}>
            <div style={{display:"inline-flex",alignItems:"center",gap:"10px",marginBottom:"6px"}}>
              <div style={{width:"36px",height:"36px",borderRadius:"8px",background:`linear-gradient(145deg,${T.gold},#A8843A)`,display:"flex",alignItems:"center",justifyContent:"center"}}>
                <Icon name="logo" size={16} color="#fff" strokeWidth={2.5}/>
              </div>
              <div style={{textAlign:"left"}}>
                <div style={{color:"#fff",fontWeight:"800",fontSize:"14px",letterSpacing:"-0.5px",lineHeight:1.2}}>고객 전용<br/>재무정보 조회 시스템</div>
                <div style={{color:T.gold,fontSize:"10px",fontWeight:"600",letterSpacing:"1px",textTransform:"uppercase"}}>Financial Portal</div>
              </div>
            </div>
          </div>

          <div style={{
            background:"rgba(255,255,255,0.97)",borderRadius:T.radiusXl,
            padding:isMobile?"28px 24px":"40px 36px",
            boxShadow:"0 24px 64px rgba(0,0,0,0.3)",
          }}>
            <h3 style={{color:T.navy,fontSize:"18px",fontWeight:"800",margin:"0 0 4px",letterSpacing:"-0.6px"}}>로그인</h3>
            <p style={{color:T.textMuted,fontSize:"12px",margin:"0 0 22px"}}>등록된 계정으로 로그인하세요</p>

            <div style={{marginBottom:"14px"}}><Label>이메일</Label><Input value={e} onChange={setE} placeholder="example@email.com" type="email" onKeyDown={ev=>ev.key==="Enter"&&go()}/></div>
            <div style={{marginBottom:"12px"}}><Label>비밀번호</Label><Input value={p} onChange={setP} placeholder="비밀번호 입력" type="password" onKeyDown={ev=>ev.key==="Enter"&&go()}/></div>

            {/* 이메일 저장 + 비밀번호 찾기 */}
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"18px"}}>
              <label style={{display:"flex",alignItems:"center",gap:"7px",cursor:"pointer",userSelect:"none"}}>
                <input type="checkbox" checked={saveEmail} onChange={ev=>setSaveEmail(ev.target.checked)}
                  style={{width:"15px",height:"15px",accentColor:T.blue,cursor:"pointer"}}/>
                <span style={{color:T.textSub,fontSize:"13px"}}>이메일 저장</span>
              </label>
              <button onClick={()=>setShowForgot(true)} style={{background:"none",border:"none",color:T.blue,fontSize:"13px",cursor:"pointer",textDecoration:"underline",textUnderlineOffset:"2px",fontFamily:T.font}}>비밀번호 찾기</button>
            </div>

            {err&&<div style={{background:T.redLight,border:`1px solid rgba(192,57,43,0.2)`,borderRadius:T.radiusSm,padding:"10px 13px",marginBottom:"16px",color:T.red,fontSize:"13px",fontWeight:"500"}}>{err}</div>}

            <button onClick={go} disabled={loading} style={{
              width:"100%",padding:"12px",borderRadius:T.radiusSm,border:"none",
              background:`linear-gradient(135deg,${T.bgDeep},${T.bgDeepAlt})`,
              color:"#fff",fontSize:"14px",fontWeight:"700",cursor:loading?"not-allowed":"pointer",
              letterSpacing:"-0.2px",fontFamily:T.font,
              boxShadow:`0 4px 16px rgba(27,42,74,0.3)`,
              marginBottom:"14px",opacity:loading?0.7:1,
            }}>{loading?"확인 중…":"로그인"}</button>

            <div style={{textAlign:"center"}}>
              <span style={{color:T.textMuted,fontSize:"13px"}}>계정이 없으신가요? </span>
              <button onClick={()=>onGo("register")} style={{background:"none",border:"none",color:T.navy,fontSize:"13px",fontWeight:"700",cursor:"pointer",textDecoration:"underline",textUnderlineOffset:"2px"}}>회원가입 신청</button>
            </div>
          </div>

          <div style={{marginTop:"14px",padding:"12px 16px",borderRadius:T.radius,background:"rgba(0,0,0,0.2)",border:"1px solid rgba(255,255,255,0.08)"}}>
            <p style={{color:"rgba(255,255,255,0.3)",fontSize:"10px",fontWeight:"600",margin:"0 0 4px",textTransform:"uppercase",letterSpacing:"0.5px"}}>테스트 계정</p>
            <p style={{color:"rgba(255,255,255,0.4)",fontSize:"11px",margin:"0 0 2px"}}>고객: hong@example.com / 1234</p>
            <p style={{color:"rgba(255,255,255,0.4)",fontSize:"11px",margin:0}}>관리자: admin@cpa.com / admin1234</p>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────
   회원가입
───────────────────────────────────────── */
function Register({ onGo }) {
  const [f,setF]=useState({name:"",email:"",pw:"",pw2:"",phone:"",memo:""});
  const [bizs,setBizs]=useState([{no:"",name:""}]);
  const [err,setErr]=useState(""), [ok,setOk]=useState(false), [loading,setL]=useState(false);
  const u=(k,v)=>setF(x=>({...x,[k]:v}));
  const submit=()=>{
    setErr("");
    if(!f.name||!f.email||!f.pw){setErr("필수 항목을 모두 입력해주세요.");return;}
    if(f.pw!==f.pw2){setErr("비밀번호가 일치하지 않습니다.");return;}
    if(f.pw.length<6){setErr("비밀번호는 6자 이상이어야 합니다.");return;}
    const vb=bizs.filter(b=>b.no.trim());
    if(!vb.length){setErr("사업자번호를 1개 이상 입력해주세요.");return;}
    setL(true);
    const bizNames = {};
    vb.forEach(b=>{ if(b.name.trim()) bizNames[b.no.trim()]=b.name.trim(); });
    db.register({name:f.name,email:f.email,password:f.pw,phone:f.phone,memo:f.memo,businesses:vb.map(b=>b.no.trim()),bizNames})
      .then(r=>{ setL(false); if(r.ok) setOk(true); else setErr(r.msg); })
      .catch(()=>{ setL(false); setErr("서버 연결 오류가 발생했습니다."); });
  };
  if(ok) return (
    <div style={{minHeight:"100vh",background:T.bg,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:T.font}}>
      <div style={{textAlign:"center",padding:"40px"}}>
        <div style={{display:"flex",justifyContent:"center",marginBottom:"20px"}}>
          <div style={{width:"64px",height:"64px",borderRadius:"50%",background:T.greenLight,display:"flex",alignItems:"center",justifyContent:"center"}}>
            <Icon name="check" size={32} color={T.green} strokeWidth={2}/>
          </div>
        </div>
        <h2 style={{color:T.text,fontSize:"22px",fontWeight:"700",margin:"0 0 10px",letterSpacing:"-0.7px"}}>가입 신청 완료</h2>
        <p style={{color:T.textSub,fontSize:"15px",lineHeight:"1.6",marginBottom:"28px"}}>검토 후 승인 안내 드리겠습니다.</p>
        <Btn onClick={()=>onGo("login")} size="lg">로그인으로 돌아가기</Btn>
      </div>
    </div>
  );
  return (
    <div style={{minHeight:"100vh",background:T.bg,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:T.font,padding:"40px 24px"}}>
      <div style={{width:"100%",maxWidth:"520px"}}>
        <div style={{textAlign:"center",marginBottom:"32px"}}>
          <h1 style={{color:T.text,fontSize:"26px",fontWeight:"700",margin:"0 0 6px",letterSpacing:"-0.8px"}}>회원가입 신청</h1>
          <p style={{color:T.textSub,fontSize:"14px",margin:0}}>승인 후 이용 가능합니다</p>
        </div>
        <Card style={{padding:"36px"}}>
          <p style={{color:T.textSub,fontSize:"11px",fontWeight:"700",textTransform:"uppercase",letterSpacing:"0.8px",margin:"0 0 16px"}}>기본 정보</p>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px",marginBottom:"12px"}}>
            <div><Label>이름 *</Label><Input value={f.name} onChange={v=>u("name",v)} placeholder="홍길동"/></div>
            <div><Label>연락처</Label><Input value={f.phone} onChange={v=>u("phone",v)} placeholder="010-0000-0000"/></div>
          </div>
          <div style={{marginBottom:"12px"}}><Label>이메일 *</Label><Input value={f.email} onChange={v=>u("email",v)} type="email" placeholder="example@email.com"/></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px",marginBottom:"28px"}}>
            <div><Label>비밀번호 * (6자 이상)</Label><Input value={f.pw} onChange={v=>u("pw",v)} type="password" placeholder="••••••"/></div>
            <div><Label>비밀번호 확인 *</Label><Input value={f.pw2} onChange={v=>u("pw2",v)} type="password" placeholder="••••••"/></div>
          </div>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"12px"}}>
            <p style={{color:T.textSub,fontSize:"11px",fontWeight:"700",textTransform:"uppercase",letterSpacing:"0.8px",margin:0}}>사업자 정보 *</p>
            <Btn onClick={()=>setBizs(b=>[...b,{no:"",name:""}])} variant="ghost" size="sm">+ 사업자 추가</Btn>
          </div>
          {bizs.map((biz,i)=>(
            <div key={i} style={{display:"flex",gap:"8px",alignItems:"flex-end",marginBottom:"8px"}}>
              <div style={{flex:1}}><Label>사업자번호</Label><Input value={biz.no} onChange={v=>{const nb=[...bizs];nb[i].no=v;setBizs(nb)}} placeholder="000-00-00000"/></div>
              <div style={{flex:1}}><Label>상호 (선택)</Label><Input value={biz.name} onChange={v=>{const nb=[...bizs];nb[i].name=v;setBizs(nb)}} placeholder="(주)한국상사"/></div>
              {bizs.length>1&&<button onClick={()=>setBizs(b=>b.filter((_,idx)=>idx!==i))} style={{background:T.redLight,border:"none",color:T.red,padding:"10px 12px",borderRadius:T.radiusSm,cursor:"pointer",fontWeight:"700",marginBottom:"1px",fontSize:"14px",lineHeight:1}}>×</button>}
            </div>
          ))}
          <div style={{marginTop:"20px",marginBottom:"24px"}}><Label>요청사항 (선택)</Label>
            <textarea value={f.memo} onChange={e=>u("memo",e.target.value)} placeholder="추가 요청사항"
              style={{width:"100%",padding:"10px 14px",borderRadius:T.radiusSm,border:`1.5px solid ${T.border}`,background:"rgba(255,255,255,0.9)",color:T.text,fontSize:"14px",fontFamily:T.font,outline:"none",resize:"none",height:"72px",boxSizing:"border-box"}}/>
          </div>
          {err&&<div style={{background:T.redLight,border:`1px solid rgba(255,59,48,0.2)`,borderRadius:T.radiusSm,padding:"11px 14px",marginBottom:"16px",color:T.red,fontSize:"13px",fontWeight:"500"}}>{err}</div>}
          <Btn onClick={submit} disabled={loading} style={{width:"100%",marginBottom:"12px"}} size="lg">{loading?"신청 중…":"가입 신청하기"}</Btn>
          <div style={{textAlign:"center"}}><button onClick={()=>onGo("login")} style={{background:"none",border:"none",color:T.textSub,fontSize:"13px",cursor:"pointer"}}>← 로그인으로 돌아가기</button></div>
        </Card>
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────
   재무제표 뷰어들
───────────────────────────────────────── */
const Empty = () => (
  <Card style={{padding:"60px",textAlign:"center"}}>
    <div style={{display:"flex",justifyContent:"center",marginBottom:"16px",opacity:0.25}}>
      <Icon name="inbox" size={48} color={T.navy} strokeWidth={1}/>
    </div>
    <p style={{color:T.textSub,fontSize:"15px",margin:"0 0 6px",fontWeight:"500"}}>데이터가 없습니다</p>
    <p style={{color:T.textMuted,fontSize:"13px",margin:0}}>담당자에게 문의해주세요.</p>
  </Card>
);

function FSTable({ data, isMobile }) {
  if(!data?.rows) return <Empty/>;

  const typeStyle = {
    total:   {fontWeight:"700",color:T.navy,background:"rgba(27,42,74,0.05)",borderTop:`1px solid rgba(27,42,74,0.2)`,borderBottom:`1px solid rgba(27,42,74,0.2)`},
    subtotal:{fontWeight:"600",color:T.text,background:"rgba(0,0,0,0.02)"},
    header:  {fontWeight:"600",color:T.text,background:"transparent"},
    item:    {fontWeight:"400",color:T.textSub,background:"transparent"},
  };

  // ── 모바일: 천원 단위 절사, 4컬럼 압축 레이아웃
  if (isMobile) {
    return (
      <div style={{...CARD_STYLE,overflow:"hidden"}}>
        {/* 단위 표시 배너 */}
        <div style={{
          padding:"7px 14px",
          background:T.bgDeep,
          display:"flex",alignItems:"center",justifyContent:"space-between",
        }}>
          <span style={{fontSize:"10px",fontWeight:"700",color:"rgba(255,255,255,0.5)",letterSpacing:"0.5px",textTransform:"uppercase"}}>재무제표</span>
          <span style={{fontSize:"10px",fontWeight:"700",color:T.gold,letterSpacing:"0.5px"}}>(단위: 천원)</span>
        </div>
        {/* 컬럼 헤더 */}
        <div style={{
          display:"grid",gridTemplateColumns:"minmax(0,1fr) 62px 62px 54px",
          padding:"8px 12px",borderBottom:`1px solid ${T.borderStrong}`,
          background:"rgba(27,42,74,0.04)",
        }}>
          {["계정과목","당기","전기","증감"].map((h,i)=>(
            <span key={h} style={{
              fontSize:"10px",fontWeight:"700",color:T.textMuted,
              textAlign:i===0?"left":"right",textTransform:"uppercase",letterSpacing:"0.5px",
            }}>{h}</span>
          ))}
        </div>

        {data.rows.map((row,i)=>{
          const st       = typeStyle[row.type]||typeStyle.item;
          const diff     = row.당기 - row.전기;
          const isTotal  = row.type==="total";
          const isHeader = row.type==="header";
          const indent   = row.level * 8;
          const diffColor= diff>0?T.green:diff<0?T.red:T.textMuted;

          return (
            <div key={i} style={{
              display:"grid",
              gridTemplateColumns:"minmax(0,1fr) 62px 62px 54px",
              padding:`${isTotal?"10px":"7px"} 12px`,
              borderBottom:`1px solid ${T.border}`,
              background:isTotal?"rgba(27,42,74,0.04)":st.background||"transparent",
              borderTop:isTotal?`1px solid rgba(27,42,74,0.12)`:"none",
              alignItems:"center",
            }}>
              <span style={{
                paddingLeft:`${indent}px`,
                fontSize:isTotal?"12px":"11px",fontWeight:st.fontWeight,
                color:isHeader?T.navy:isTotal?T.navy:st.color,
                lineHeight:"1.3",letterSpacing:"-0.1px",
                overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",
              }}>
                {isHeader&&<span style={{marginRight:"5px",display:"inline-flex",verticalAlign:"middle"}}><Icon name="dot" size={6} color={T.gold}/></span>}
                {row.계정과목}
              </span>
              <span style={{textAlign:"right",fontSize:isTotal?"12px":"11px",fontWeight:st.fontWeight,color:isTotal?T.navy:T.text,fontVariantNumeric:"tabular-nums"}}>
                {FK(row.당기)}
              </span>
              <span style={{textAlign:"right",fontSize:"11px",color:T.textMuted,fontVariantNumeric:"tabular-nums"}}>
                {FK(row.전기)}
              </span>
              <span style={{textAlign:"right",fontSize:"10px",fontWeight:"600",color:diffColor,fontVariantNumeric:"tabular-nums"}}>
                {diff===0?"-":(diff>0?"▲":"▼")+FK(Math.abs(diff))}
              </span>
            </div>
          );
        })}
      </div>
    );
  }

  // ── PC: 천원 단위 테이블
  return (
    <div style={{...CARD_STYLE,overflow:"hidden"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:"13px",fontFamily:T.font,tableLayout:"fixed"}}>
        <thead>
          <tr style={{background:T.bgDeep}}>
            <th style={{padding:"12px 16px",textAlign:"left",fontSize:"11px",fontWeight:"700",color:"rgba(255,255,255,0.7)",letterSpacing:"0.4px",textTransform:"uppercase"}}>
              계정과목
            </th>
            <th style={{padding:"12px 16px",textAlign:"right",fontSize:"11px",fontWeight:"700",color:"rgba(255,255,255,0.7)",letterSpacing:"0.4px",textTransform:"uppercase",width:"120px"}}>
              당기
            </th>
            <th style={{padding:"12px 16px",textAlign:"right",fontSize:"11px",fontWeight:"700",color:"rgba(255,255,255,0.7)",letterSpacing:"0.4px",textTransform:"uppercase",width:"120px"}}>
              전기
            </th>
            <th style={{padding:"12px 16px",textAlign:"right",fontSize:"11px",fontWeight:"700",color:"rgba(255,255,255,0.7)",letterSpacing:"0.4px",textTransform:"uppercase",width:"100px"}}>
              증감
            </th>
            <th style={{padding:"12px 16px",textAlign:"right",fontSize:"10px",fontWeight:"600",color:T.gold,letterSpacing:"0.3px",width:"90px",whiteSpace:"nowrap"}}>
              단위: 천원
            </th>
          </tr>
        </thead>
        <tbody>
          {data.rows.map((row,i)=>{
            const st=typeStyle[row.type]||typeStyle.item;
            const diff=row.당기-row.전기;
            const diffColor=diff>0?T.green:diff<0?T.red:T.textMuted;
            return (
              <tr key={i} style={{borderBottom:`1px solid ${T.border}`,...st}}>
                <td style={{padding:`10px 16px 10px ${16+row.level*16}px`,...st,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                  {row.type==="header"&&<span style={{marginRight:"6px",display:"inline-flex",verticalAlign:"middle"}}><Icon name="dot" size={6} color={T.gold}/></span>}
                  {row.계정과목}
                </td>
                <td style={{padding:"10px 16px",textAlign:"right",fontVariantNumeric:"tabular-nums",...st}}>{FK(row.당기)}</td>
                <td style={{padding:"10px 16px",textAlign:"right",fontVariantNumeric:"tabular-nums",color:T.textMuted}}>{FK(row.전기)}</td>
                <td style={{padding:"10px 16px",textAlign:"right",fontVariantNumeric:"tabular-nums",color:diffColor,fontSize:"12px"}}>
                  {diff===0?"-":(diff>0?"▲ ":"▼ ")+FK(Math.abs(diff))}
                </td>
                <td/>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function BSView({ data, prevData, prevYear, isMobile }) {
  if(!data?.rows) return <Empty/>;
  const get = (d, kw) => d?.rows?.find(r=>r.계정과목.includes(kw))?.당기||0;
  const 자산=get(data,"자산총계"), 부채=get(data,"부채총계"), 자본=get(data,"자본총계");
  const p자산=get(prevData,"자산총계"), p부채=get(prevData,"부채총계"), p자본=get(prevData,"자본총계");
  const 부채비율 = 자본?((부채/자본)*100).toFixed(1):"-";
  const p부채비율 = p자본?((p부채/p자본)*100).toFixed(1):null;

  const hasPrev = prevData?.rows && p자산 > 0;

  const compareData = [
    { name:"자산총계", 당기:자산, 전기:p자산 },
    { name:"부채총계", 당기:부채, 전기:p부채 },
    { name:"자본총계", 당기:자본, 전기:p자본 },
  ];

  const yFmt = v => v>=100000000?`${(v/100000000).toFixed(0)}억`:`${(v/10000).toFixed(0)}만`;
  const pctChg = (cur,prev) => prev ? ((cur-prev)/Math.abs(prev)*100).toFixed(1) : null;

  return (
    <div>
      {/* KPI */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:"12px",marginBottom:"20px"}}>
        {[
          {label:"자산총계", cur:자산, prev:p자산, color:T.blue},
          {label:"부채총계", cur:부채, prev:p부채, color:T.red},
          {label:"자본총계", cur:자본, prev:p자본, color:T.green},
          {label:"부채비율",  cur:부채비율+"%", prev:p부채비율?p부채비율+"%":null, color:T.orange, raw:true},
        ].map(k=>{
          const chg = !k.raw ? pctChg(k.cur, k.prev) : null;
          const up = chg > 0;
          return (
            <Card key={k.label} style={{padding:"18px 20px"}}>
              <p style={{fontSize:"11px",fontWeight:"600",color:T.textSub,margin:"0 0 10px",letterSpacing:"-0.1px"}}>{k.label}</p>
              <p style={{fontSize:"20px",fontWeight:"700",color:T.text,margin:"0 0 6px",fontVariantNumeric:"tabular-nums",letterSpacing:"-0.8px"}}>
                {k.raw ? k.cur : FW(k.cur)}
              </p>
              {hasPrev && (
                <div style={{display:"flex",alignItems:"center",gap:"6px"}}>
                  {chg !== null && (
                    <span style={{fontSize:"11px",fontWeight:"600",color:up?T.green:T.red}}>
                      {up?"▲":"▼"} {Math.abs(chg)}%
                    </span>
                  )}
                  <span style={{fontSize:"11px",color:T.textMuted}}>
                    전년 {k.raw ? k.prev : FW(k.prev)}
                  </span>
                </div>
              )}
            </Card>
          );
        })}
      </div>

      {/* 전년도 비교 차트 */}
      {hasPrev && (
        <Card style={{padding:"24px",marginBottom:"16px"}} className="chart-area">
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"18px"}}>
            <p style={{color:T.text,fontWeight:"600",fontSize:"13px",margin:0,letterSpacing:"-0.2px"}}>재무구조 전년도 비교</p>
            <div style={{display:"flex",gap:"14px",alignItems:"center"}}>
              <span style={{display:"flex",alignItems:"center",gap:"5px",fontSize:"12px",color:T.textSub}}>
                <span style={{width:"10px",height:"10px",borderRadius:"3px",background:T.blue,display:"inline-block"}}/>당기
              </span>
              <span style={{display:"flex",alignItems:"center",gap:"5px",fontSize:"12px",color:T.textSub}}>
                <span style={{width:"10px",height:"10px",borderRadius:"3px",background:"rgba(0,113,227,0.25)",display:"inline-block"}}/>전기 ({prevYear})
              </span>
            </div>
          </div>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={compareData} barCategoryGap="30%" barGap={4}>
              <XAxis dataKey="name" tick={{fill:T.textSub,fontSize:12,fontFamily:T.font}} axisLine={false} tickLine={false}/>
              <YAxis tick={{fill:T.textSub,fontSize:10,fontFamily:T.font}} axisLine={false} tickLine={false} tickFormatter={yFmt}/>
              <Tooltip
                contentStyle={{background:"#fff",border:`1px solid ${T.border}`,borderRadius:T.radiusSm,fontFamily:T.font,fontSize:"12px",boxShadow:T.shadowMd}}
                formatter={(v,name)=>[FW(v), name==="당기"?"당기":prevYear+"년(전기)"]}
              />
              <Bar dataKey="당기"  fill={T.blue}                    radius={[5,5,0,0]} name="당기"/>
              <Bar dataKey="전기"  fill="rgba(0,113,227,0.25)"      radius={[5,5,0,0]} name="전기"/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      )}

      <FSTable data={data} isMobile={isMobile}/>
    </div>
  );
}

function ISView({ data, prevData, prevYear, isMobile }) {
  if(!data?.rows) return <Empty/>;
  const get  = (d,kw) => d?.rows?.find(r=>r.계정과목.includes(kw))?.당기||0;
  const 매출액=get(data,"매출액"), 영업이익=get(data,"영업이익"), 순이익=get(data,"당기순이익");
  const p매출=get(prevData,"매출액"), p영업=get(prevData,"영업이익"), p순이익=get(prevData,"당기순이익");
  const 영업이익률=매출액?((영업이익/매출액)*100).toFixed(1):"-";
  const 순이익률  =매출액?((순이익/매출액)*100).toFixed(1):"-";
  const p영업이익률=p매출?((p영업/p매출)*100).toFixed(1):null;
  const p순이익률  =p매출?((p순이익/p매출)*100).toFixed(1):null;
  const hasPrev = prevData?.rows && p매출 > 0;

  const pctChg=(cur,prev)=>prev?((cur-prev)/Math.abs(prev)*100).toFixed(1):null;

  // 당기·전기 나란히 비교 막대차트
  const barData=[
    {name:"매출액",    당기:매출액,   전기:p매출},
    {name:"매출총이익",당기:get(data,"매출총이익"), 전기:get(prevData,"매출총이익")},
    {name:"영업이익",  당기:영업이익, 전기:p영업},
    {name:"당기순이익",당기:순이익,   전기:p순이익},
  ];

  // 이익률 추이 꺾은선차트 (전기→당기)
  const lineData = hasPrev ? [
    {year:prevYear,  영업이익률:Number(p영업이익률), 순이익률:Number(p순이익률)},
    {year:"당기",    영업이익률:Number(영업이익률),  순이익률:Number(순이익률)},
  ] : null;

  const yFmt=v=>v>=100000000?`${(v/100000000).toFixed(0)}억`:`${(v/10000).toFixed(0)}만`;
  const Legend=({color,label})=>(
    <span style={{display:"flex",alignItems:"center",gap:"5px",fontSize:"12px",color:T.textSub}}>
      <span style={{width:"10px",height:"10px",borderRadius:"3px",background:color,display:"inline-block"}}/>
      {label}
    </span>
  );

  return (
    <div>
      {/* KPI */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:"12px",marginBottom:"20px"}}>
        {[
          {label:"매출액",    cur:매출액,   prev:p매출,   color:T.blue},
          {label:"영업이익",  cur:영업이익, prev:p영업,   color:T.orange},
          {label:"당기순이익",cur:순이익,   prev:p순이익, color:T.green},
          {label:"영업이익률",cur:영업이익률+"%", prev:p영업이익률?p영업이익률+"%":null, color:T.purple, raw:true},
          {label:"순이익률",  cur:순이익률+"%",   prev:p순이익률?p순이익률+"%":null,     color:T.green,  raw:true},
        ].map(k=>{
          const chg = !k.raw ? pctChg(k.cur,k.prev) : null;
          const up  = chg > 0;
          return (
            <Card key={k.label} style={{padding:"16px 18px"}}>
              <p style={{fontSize:"11px",fontWeight:"600",color:T.textSub,margin:"0 0 8px"}}>{k.label}</p>
              <p style={{fontSize:"18px",fontWeight:"700",color:T.text,margin:"0 0 5px",fontVariantNumeric:"tabular-nums",letterSpacing:"-0.6px"}}>
                {k.raw ? k.cur : FW(k.cur)}
              </p>
              {hasPrev && (
                <div style={{display:"flex",alignItems:"center",gap:"5px"}}>
                  {chg!==null && <span style={{fontSize:"11px",fontWeight:"600",color:up?T.green:T.red}}>{up?"▲":"▼"}{Math.abs(chg)}%</span>}
                  <span style={{fontSize:"11px",color:T.textMuted}}>전년 {k.raw?k.prev:FW(k.prev)}</span>
                </div>
              )}
            </Card>
          );
        })}
      </div>

      {/* 비교 차트 2개 나란히 */}
      {hasPrev && (
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1.4fr 1fr",gap:"14px",marginBottom:"16px"}} className="chart-area">
          <Card style={{padding:"22px"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"16px"}}>
              <p style={{color:T.text,fontWeight:"600",fontSize:"13px",margin:0,letterSpacing:"-0.2px"}}>주요 항목 전년 비교</p>
              <div style={{display:"flex",gap:"12px"}}><Legend color={T.blue} label="당기"/><Legend color="rgba(0,113,227,0.28)" label={`${prevYear}년`}/></div>
            </div>
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={barData} barCategoryGap="28%" barGap={3}>
                <XAxis dataKey="name" tick={{fill:T.textSub,fontSize:11,fontFamily:T.font}} axisLine={false} tickLine={false}/>
                <YAxis tick={{fill:T.textSub,fontSize:10,fontFamily:T.font}} axisLine={false} tickLine={false} tickFormatter={yFmt}/>
                <Tooltip contentStyle={{background:"#fff",border:`1px solid ${T.border}`,borderRadius:T.radiusSm,fontFamily:T.font,fontSize:"12px",boxShadow:T.shadowMd}}
                  formatter={(v,name)=>[FW(v),name==="당기"?"당기":prevYear+"년"]}/>
                <Bar dataKey="당기" fill={T.blue}                    radius={[4,4,0,0]} name="당기"/>
                <Bar dataKey="전기" fill="rgba(0,113,227,0.28)"      radius={[4,4,0,0]} name="전기"/>
              </BarChart>
            </ResponsiveContainer>
          </Card>
          <Card style={{padding:"22px"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"16px"}}>
              <p style={{color:T.text,fontWeight:"600",fontSize:"13px",margin:0,letterSpacing:"-0.2px"}}>이익률 추이</p>
              <div style={{display:"flex",gap:"12px"}}>
                <Legend color={T.orange} label="영업이익률"/>
                <Legend color={T.green}  label="순이익률"/>
              </div>
            </div>
            <ResponsiveContainer width="100%" height={200}>
              <LineChart data={lineData} margin={{top:10,right:10,left:0,bottom:0}}>
                <CartesianGrid strokeDasharray="3 3" stroke={T.border}/>
                <XAxis dataKey="year" tick={{fill:T.textSub,fontSize:12,fontFamily:T.font}} axisLine={false} tickLine={false}/>
                <YAxis tick={{fill:T.textSub,fontSize:10,fontFamily:T.font}} axisLine={false} tickLine={false} tickFormatter={v=>`${v}%`} domain={["auto","auto"]}/>
                <Tooltip contentStyle={{background:"#fff",border:`1px solid ${T.border}`,borderRadius:T.radiusSm,fontFamily:T.font,fontSize:"12px",boxShadow:T.shadowMd}}
                  formatter={(v,name)=>[`${v}%`,name]}/>
                <Line dataKey="영업이익률" stroke={T.orange} strokeWidth={2.5} dot={{fill:T.orange,r:5}} activeDot={{r:7}} name="영업이익률"/>
                <Line dataKey="순이익률"   stroke={T.green}  strokeWidth={2.5} dot={{fill:T.green, r:5}} activeDot={{r:7}} name="순이익률"/>
              </LineChart>
            </ResponsiveContainer>
          </Card>
        </div>
      )}

      {/* 테이블 */}
      <FSTable data={data} isMobile={isMobile}/>
    </div>
  );
}

function VATView({ data, prevData, prevYear, isMobile }) {
  if(!data?.rows) return <Empty/>;
  const tot={매출:0,매입:0,납부:0};
  data.rows.forEach(r=>{tot.매출+=r.매출세액;tot.매입+=r.매입세액;tot.납부+=r.납부세액;});
  const ptot={매출:0,매입:0,납부:0};
  prevData?.rows?.forEach(r=>{ptot.매출+=r.매출세액;ptot.매입+=r.매입세액;ptot.납부+=r.납부세액;});
  const hasPrev = prevData?.rows && ptot.납부 > 0;
  const pctChg=(c,p)=>p?((c-p)/Math.abs(p)*100).toFixed(1):null;

  // 기수별 당기·전기 납부세액 비교
  const barData = data.rows.map((r,i)=>({
    name: r.기수?.replace("기 ","기\n")||`${i+1}기`,
    당기: r.납부세액,
    전기: prevData?.rows?.[i]?.납부세액||0,
  }));

  return (
    <div>
      {/* KPI */}
      <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"repeat(3,1fr)",gap:"12px",marginBottom:"20px"}}>
        {[
          {label:"연간 매출세액",cur:tot.매출,prev:ptot.매출,color:T.blue},
          {label:"연간 매입세액",cur:tot.매입,prev:ptot.매입,color:T.textSub},
          {label:"연간 납부세액",cur:tot.납부,prev:ptot.납부,color:T.green},
        ].map(k=>{
          const chg=pctChg(k.cur,k.prev); const up=chg>0;
          return (
            <Card key={k.label} style={{padding:"18px 20px"}}>
              <p style={{fontSize:"11px",fontWeight:"600",color:T.textSub,margin:"0 0 8px"}}>{k.label}</p>
              <p style={{fontSize:"20px",fontWeight:"700",color:T.text,margin:"0 0 6px",fontVariantNumeric:"tabular-nums",letterSpacing:"-0.7px"}}>{FW(k.cur)}</p>
              {hasPrev&&<div style={{display:"flex",gap:"6px",alignItems:"center"}}>
                {chg!==null&&<span style={{fontSize:"11px",fontWeight:"600",color:up?T.green:T.red}}>{up?"▲":"▼"}{Math.abs(chg)}%</span>}
                <span style={{fontSize:"11px",color:T.textMuted}}>전년 {FW(k.prev)}</span>
              </div>}
            </Card>
          );
        })}
      </div>

      {/* 기수별 납부세액 비교 차트 */}
      {hasPrev && (
        <Card style={{padding:"22px",marginBottom:"16px"}} className="chart-area">
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"16px"}}>
            <p style={{color:T.text,fontWeight:"600",fontSize:"13px",margin:0,letterSpacing:"-0.2px"}}>기수별 납부세액 전년 비교</p>
            <div style={{display:"flex",gap:"14px"}}>
              <span style={{display:"flex",alignItems:"center",gap:"5px",fontSize:"12px",color:T.textSub}}>
                <span style={{width:"10px",height:"10px",borderRadius:"3px",background:T.blue,display:"inline-block"}}/>당기
              </span>
              <span style={{display:"flex",alignItems:"center",gap:"5px",fontSize:"12px",color:T.textSub}}>
                <span style={{width:"10px",height:"10px",borderRadius:"3px",background:"rgba(0,113,227,0.25)",display:"inline-block"}}/>{prevYear}년
              </span>
            </div>
          </div>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={barData} barCategoryGap="30%" barGap={4}>
              <XAxis dataKey="name" tick={{fill:T.textSub,fontSize:11,fontFamily:T.font}} axisLine={false} tickLine={false}/>
              <YAxis tick={{fill:T.textSub,fontSize:10,fontFamily:T.font}} axisLine={false} tickLine={false}
                tickFormatter={v=>v>=10000000?`${(v/10000000).toFixed(0)}천만`:`${(v/10000).toFixed(0)}만`}/>
              <Tooltip contentStyle={{background:"#fff",border:`1px solid ${T.border}`,borderRadius:T.radiusSm,fontFamily:T.font,fontSize:"12px",boxShadow:T.shadowMd}}
                formatter={(v,name)=>[FW(v),name==="당기"?"당기":prevYear+"년"]}/>
              <Bar dataKey="당기" fill={T.blue}                  radius={[4,4,0,0]} name="당기"/>
              <Bar dataKey="전기" fill="rgba(0,113,227,0.25)"    radius={[4,4,0,0]} name="전기"/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      )}

      {/* 테이블 */}
      <Card style={{overflow:"hidden"}}>
        <div className={isMobile?"mob-scroll":""}>
        <table style={{width:"100%",borderCollapse:"collapse",fontFamily:T.font,fontSize:"13px"}}>
          <thead><tr style={{borderBottom:`1px solid ${T.border}`,background:"rgba(0,0,0,0.02)"}}>
            {["신고기수","대상기간","신고기한","매출세액","매입세액","납부(환급)세액","상태"].map(h=>(
              <th key={h} style={{padding:"12px 16px",textAlign:h.includes("세액")?"right":"left",fontSize:"11px",fontWeight:"700",color:T.textSub,textTransform:"uppercase",letterSpacing:"0.4px"}}>{h}</th>
            ))}
          </tr></thead>
          <tbody>
            {data.rows.map((r,i)=>(
              <tr key={i} style={{borderBottom:`1px solid ${T.border}`}}>
                <td style={{padding:"13px 16px",fontWeight:"600",color:T.text}}>{r.기수}</td>
                <td style={{padding:"13px 16px",color:T.textSub}}>{r.기간}</td>
                <td style={{padding:"13px 16px",color:T.textSub}}>{r.신고기한}</td>
                <td style={{padding:"13px 16px",textAlign:"right",fontVariantNumeric:"tabular-nums",color:T.text}}>{F(r.매출세액)}</td>
                <td style={{padding:"13px 16px",textAlign:"right",fontVariantNumeric:"tabular-nums",color:T.textSub}}>{F(r.매입세액)}</td>
                <td style={{padding:"13px 16px",textAlign:"right",fontVariantNumeric:"tabular-nums",fontWeight:"600",color:T.blue}}>{F(r.납부세액)}</td>
                <td style={{padding:"13px 16px"}}><Pill label={r.상태||"완료"} color={T.green} bg={T.greenLight}/></td>
              </tr>
            ))}
          </tbody>
          <tfoot><tr style={{borderTop:`1.5px solid ${T.borderStrong}`,background:"rgba(0,113,227,0.03)"}}>
            <td colSpan="3" style={{padding:"12px 16px",fontWeight:"700",color:T.text,fontSize:"13px"}}>합계</td>
            <td style={{padding:"12px 16px",textAlign:"right",fontWeight:"700",fontVariantNumeric:"tabular-nums",color:T.text}}>{F(tot.매출)}</td>
            <td style={{padding:"12px 16px",textAlign:"right",fontWeight:"700",fontVariantNumeric:"tabular-nums",color:T.textSub}}>{F(tot.매입)}</td>
            <td style={{padding:"12px 16px",textAlign:"right",fontWeight:"700",fontVariantNumeric:"tabular-nums",color:T.blue}}>{F(tot.납부)}</td>
            <td/>
          </tr></tfoot>
        </table>
        </div>
      </Card>
    </div>
  );
}

function TaxView({ data, prevData, prevYear, isMobile }) {
  if(!data) return <Empty/>;
  const eff=data.과세표준?((data.산출세액/data.과세표준)*100).toFixed(2):"-";
  const pEff=prevData?.과세표준?((prevData.산출세액/prevData.과세표준)*100).toFixed(2):null;
  const hasPrev = prevData && prevData.납부세액 > 0;
  const pctChg=(c,p)=>p?((c-p)/Math.abs(p)*100).toFixed(1):null;

  // 당기·전기 비교 막대
  const barData=[
    {name:"과세표준", 당기:data.과세표준,    전기:prevData?.과세표준||0},
    {name:"산출세액", 당기:data.산출세액,    전기:prevData?.산출세액||0},
    {name:"납부세액", 당기:data.납부세액,    전기:prevData?.납부세액||0},
  ];

  const rows=[
    ["사업연도",data.사업연도],["신고기한",data.신고기한],
    ["과세표준",FW(data.과세표준)],["산출세액",FW(data.산출세액)],
    ["공제·감면세액",FW(data.공제감면세액)],["최종 납부세액",FW(data.납부세액)],
  ];

  return (
    <div>
      {/* KPI */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:"12px",marginBottom:"20px"}}>
        {[
          {label:"과세표준",    cur:data.과세표준,     prev:prevData?.과세표준,    color:T.textSub},
          {label:"산출세액",    cur:data.산출세액,     prev:prevData?.산출세액,    color:T.blue},
          {label:"공제·감면",  cur:data.공제감면세액, prev:prevData?.공제감면세액,color:T.green},
          {label:"최종납부세액",cur:data.납부세액,     prev:prevData?.납부세액,    color:T.orange},
          {label:"실효세율",    cur:eff+"%",           prev:pEff?pEff+"%":null,    color:T.purple, raw:true},
        ].map(k=>{
          const chg=!k.raw?pctChg(k.cur,k.prev):null; const up=chg>0;
          return (
            <Card key={k.label} style={{padding:"16px 18px"}}>
              <p style={{fontSize:"11px",fontWeight:"600",color:T.textSub,margin:"0 0 8px"}}>{k.label}</p>
              <p style={{fontSize:"16px",fontWeight:"700",color:T.text,margin:"0 0 5px",fontVariantNumeric:"tabular-nums",letterSpacing:"-0.5px"}}>
                {k.raw?k.cur:FW(k.cur)}
              </p>
              {hasPrev&&<div style={{display:"flex",gap:"5px",alignItems:"center"}}>
                {chg!==null&&<span style={{fontSize:"11px",fontWeight:"600",color:up?T.green:T.red}}>{up?"▲":"▼"}{Math.abs(chg)}%</span>}
                <span style={{fontSize:"11px",color:T.textMuted}}>전년 {k.raw?k.prev:FW(k.prev)}</span>
              </div>}
            </Card>
          );
        })}
      </div>

      <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:"16px"}} className="two-col">
        <Card style={{overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontFamily:T.font,fontSize:"13px"}}>
            {rows.map(([k,v],i)=>(
              <tr key={i} style={{borderBottom:`1px solid ${T.border}`}}>
                <td style={{padding:"13px 20px",color:T.textSub,fontWeight:"500"}}>{k}</td>
                <td style={{padding:"13px 20px",textAlign:"right",fontWeight:"600",color:T.text,fontVariantNumeric:"tabular-nums"}}>{v||"-"}</td>
              </tr>
            ))}
          </table>
        </Card>
        <div>
          {/* 전년도 비교 막대차트 */}
          {hasPrev ? (
            <Card style={{padding:"20px",marginBottom:"12px"}} className="chart-area">
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"14px"}}>
                <p style={{color:T.text,fontWeight:"600",fontSize:"13px",margin:0,letterSpacing:"-0.2px"}}>전년도 비교</p>
                <div style={{display:"flex",gap:"10px"}}>
                  <span style={{display:"flex",alignItems:"center",gap:"4px",fontSize:"11px",color:T.textSub}}>
                    <span style={{width:"8px",height:"8px",borderRadius:"2px",background:T.blue,display:"inline-block"}}/>당기
                  </span>
                  <span style={{display:"flex",alignItems:"center",gap:"4px",fontSize:"11px",color:T.textSub}}>
                    <span style={{width:"8px",height:"8px",borderRadius:"2px",background:"rgba(0,113,227,0.28)",display:"inline-block"}}/>{prevYear}
                  </span>
                </div>
              </div>
              <ResponsiveContainer width="100%" height={160}>
                <BarChart data={barData} barCategoryGap="32%" barGap={3}>
                  <XAxis dataKey="name" tick={{fill:T.textSub,fontSize:11,fontFamily:T.font}} axisLine={false} tickLine={false}/>
                  <YAxis tick={{fill:T.textSub,fontSize:10,fontFamily:T.font}} axisLine={false} tickLine={false}
                    tickFormatter={v=>v>=100000000?`${(v/100000000).toFixed(0)}억`:`${(v/10000).toFixed(0)}만`}/>
                  <Tooltip contentStyle={{background:"#fff",border:`1px solid ${T.border}`,borderRadius:T.radiusSm,fontFamily:T.font,fontSize:"12px",boxShadow:T.shadowMd}}
                    formatter={(v,name)=>[FW(v),name==="당기"?"당기":prevYear+"년"]}/>
                  <Bar dataKey="당기" fill={T.blue}                  radius={[4,4,0,0]} name="당기"/>
                  <Bar dataKey="전기" fill="rgba(0,113,227,0.28)"    radius={[4,4,0,0]} name="전기"/>
                </BarChart>
              </ResponsiveContainer>
            </Card>
          ) : (
            <Card style={{padding:"28px",textAlign:"center",marginBottom:"12px",background:data.신고상태==="완료"?T.greenLight:`rgba(255,149,0,0.08)`,border:`1px solid ${data.신고상태==="완료"?"rgba(52,199,89,0.25)":"rgba(255,149,0,0.25)"}`}}>
              <div style={{display:"flex",justifyContent:"center",marginBottom:"10px"}}>
                <div style={{width:"52px",height:"52px",borderRadius:"50%",background:data.신고상태==="완료"?T.greenLight:T.orangeLight,display:"flex",alignItems:"center",justifyContent:"center"}}>
                  <Icon name={data.신고상태==="완료"?"check":"clock"} size={26} color={data.신고상태==="완료"?T.green:T.orange} strokeWidth={2}/>
                </div>
              </div>
              <p style={{color:data.신고상태==="완료"?T.green:T.orange,fontSize:"17px",fontWeight:"700",margin:"0 0 6px",letterSpacing:"-0.5px"}}>신고 {data.신고상태}</p>
              <p style={{color:T.textSub,fontSize:"13px",margin:0}}>신고기한: {data.신고기한}</p>
            </Card>
          )}
          {hasPrev && (
            <Card style={{padding:"20px",marginBottom:"12px",textAlign:"center",background:data.신고상태==="완료"?T.greenLight:`rgba(255,149,0,0.08)`,border:`1px solid ${data.신고상태==="완료"?"rgba(52,199,89,0.25)":"rgba(255,149,0,0.25)"}`}}>
              <div style={{display:"flex",justifyContent:"center",marginBottom:"8px"}}>
                <div style={{width:"42px",height:"42px",borderRadius:"50%",background:data.신고상태==="완료"?T.greenLight:T.orangeLight,display:"flex",alignItems:"center",justifyContent:"center"}}>
                  <Icon name={data.신고상태==="완료"?"check":"clock"} size={22} color={data.신고상태==="완료"?T.green:T.orange} strokeWidth={2}/>
                </div>
              </div>
              <p style={{color:data.신고상태==="완료"?T.green:T.orange,fontSize:"15px",fontWeight:"700",margin:"0 0 4px"}}>신고 {data.신고상태}</p>
              <p style={{color:T.textSub,fontSize:"12px",margin:0}}>신고기한: {data.신고기한}</p>
            </Card>
          )}
          <Card style={{padding:"16px 18px"}}>
            <p style={{color:T.orange,fontSize:"12px",fontWeight:"600",margin:"0 0 8px"}}>법인세 세율 (2024년 기준)</p>
            {[["2억 이하","9%"],["2억 ~ 200억","19%"],["200억 초과","21%"],["3,000억 초과","24%"]].map(([k,v])=>(
              <div key={k} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:`1px solid ${T.border}`}}>
                <span style={{color:T.textSub,fontSize:"12px"}}>{k}</span>
                <span style={{color:T.text,fontSize:"12px",fontWeight:"600"}}>{v}</span>
              </div>
            ))}
            <p style={{color:T.textMuted,fontSize:"11px",margin:"8px 0 0"}}>지방소득세 10% 별도 납부</p>
          </Card>
        </div>
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────
   Excel 업로드 드롭존
───────────────────────────────────────── */
function ExcelDropzone({ onParsed, docType }) {
  const [drag,setDrag]=useState(false);
  const [status,setStatus]=useState("idle"); // idle | parsing | ok | error
  const [msg,setMsg]=useState("");
  const ref=useRef();

  const process = useCallback((file) => {
    if(!file) return;
    if(!file.name.match(/\.(xlsx|xls)$/i)){setStatus("error");setMsg("Excel 파일(.xlsx, .xls)만 업로드 가능합니다.");return;}
    setStatus("parsing"); setMsg("");
    const reader=new FileReader();
    reader.onload=(e)=>{
      try {
        const result=parseExcel(e.target.result, docType);
        if(!result){setStatus("error");setMsg("파일을 파싱할 수 없습니다. 위하고T 출력 형식인지 확인해주세요.");return;}
        setStatus("ok"); setMsg(file.name);
        onParsed(result);
      } catch(err) {
        setStatus("error"); setMsg("파싱 오류: "+err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  },[docType,onParsed]);

  const onDrop=(e)=>{ e.preventDefault(); setDrag(false); process(e.dataTransfer.files[0]); };

  return (
    <div>
      <div
        onClick={()=>ref.current?.click()}
        onDragOver={e=>{e.preventDefault();setDrag(true);}}
        onDragLeave={()=>setDrag(false)}
        onDrop={onDrop}
        style={{
          border:`2px dashed ${drag?T.blue:status==="ok"?T.green:status==="error"?T.red:T.border}`,
          borderRadius:T.radius, padding:"36px 24px", textAlign:"center", cursor:"pointer",
          background:drag?T.blueLight:status==="ok"?T.greenLight:status==="error"?T.redLight:"rgba(0,0,0,0.015)",
          transition:"all 0.2s ease",
        }}
      >
        <input ref={ref} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>process(e.target.files[0])}/>
        <div style={{display:"flex",justifyContent:"center",marginBottom:"12px"}}>
          <div style={{width:"48px",height:"48px",borderRadius:"12px",
            background:status==="ok"?T.greenLight:status==="error"?T.redLight:status==="parsing"?T.blueLight:"rgba(0,0,0,0.04)",
            display:"flex",alignItems:"center",justifyContent:"center",
          }}>
            <Icon name={status==="parsing"?"clock":status==="ok"?"check":status==="error"?"alert":"upload"}
              size={24}
              color={status==="ok"?T.green:status==="error"?T.red:status==="parsing"?T.blue:T.textMuted}
              strokeWidth={1.6}/>
          </div>
        </div>
        <p style={{color:T.text,fontWeight:"600",fontSize:"14px",margin:"0 0 4px",letterSpacing:"-0.3px"}}>
          {status==="parsing"?"파싱 중…":status==="ok"?"업로드 완료":status==="error"?"업로드 실패":"Excel 파일을 드래그하거나 클릭하여 업로드"}
        </p>
        <p style={{color:status==="error"?T.red:T.textSub,fontSize:"13px",margin:0}}>
          {msg||"위하고T 엑셀 변환 파일을 지원합니다 (.xlsx, .xls)"}
        </p>
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────
   세금 납부 일정
   근거: 국세기본법·부가가치세법·법인세법·소득세법·
         국민연금법·건강보험법·고용보험법·산재보험법
───────────────────────────────────────── */

// 연도·사업자유형에 따른 전체 세금 일정 생성
const buildSchedule = (year, bizType) => {
  const y = Number(year);
  const isCorpBase = bizType === "법인";

  const items = [
    // ── 부가가치세 (부가가치세법 제48·49조) ──────────────────
    { id:"vat1", cat:"부가가치세", label:"1기 예정신고·납부", date:`${y}-04-25`, type:"VAT",
      desc:"1~3월분 과세기간 예정신고", law:"부가가치세법 제48조", color:"#0071E3" },
    { id:"vat2", cat:"부가가치세", label:"1기 확정신고·납부", date:`${y}-07-25`, type:"VAT",
      desc:"4~6월분 포함 1기 전체 확정신고", law:"부가가치세법 제49조", color:"#0071E3" },
    { id:"vat3", cat:"부가가치세", label:"2기 예정신고·납부", date:`${y}-10-25`, type:"VAT",
      desc:"7~9월분 과세기간 예정신고", law:"부가가치세법 제48조", color:"#0071E3" },
    { id:"vat4", cat:"부가가치세", label:"2기 확정신고·납부", date:`${y+1}-01-25`, type:"VAT",
      desc:"10~12월분 포함 2기 전체 확정신고", law:"부가가치세법 제49조", color:"#0071E3" },

    // ── 법인세 (법인세법 제60조) ─────────────────────────────
    ...(isCorpBase ? [
      { id:"cit1", cat:"법인세", label:"법인세 중간예납", date:`${y}-08-31`, type:"CIT",
        desc:`${y-1}년 산출세액 기준 50% 납부`, law:"법인세법 제63조", color:"#5E5CE6" },
      { id:"cit2", cat:"법인세", label:"법인세 신고·납부", date:`${y+1}-03-31`, type:"CIT",
        desc:`${y}년 사업연도 법인세 확정신고`, law:"법인세법 제60조", color:"#5E5CE6" },
    ] : [
    // ── 종합소득세 (소득세법 제70조) ────────────────────────
      { id:"pit1", cat:"종합소득세", label:"종합소득세 중간예납", date:`${y}-11-30`, type:"PIT",
        desc:`전년도 종합소득세 50% 중간예납`, law:"소득세법 제65조", color:"#AF52DE" },
      { id:"pit2", cat:"종합소득세", label:"종합소득세 신고·납부", date:`${y+1}-05-31`, type:"PIT",
        desc:`${y}년 귀속 종합소득세 확정신고`, law:"소득세법 제70조", color:"#AF52DE" },
    ]),

    // ── 원천세 (소득세법 제128조) ────────────────────────────
    { id:"wt1",  cat:"원천세", label:"원천세 납부 (1월분)",  date:`${y}-02-10`, type:"WHT", desc:"1월 급여 원천징수세액 납부", law:"소득세법 제128조", color:"#FF9500" },
    { id:"wt2",  cat:"원천세", label:"원천세 납부 (2월분)",  date:`${y}-03-10`, type:"WHT", desc:"2월 급여 원천징수세액 납부", law:"소득세법 제128조", color:"#FF9500" },
    { id:"wt3",  cat:"원천세", label:"원천세 납부 (3월분)",  date:`${y}-04-10`, type:"WHT", desc:"3월 급여 원천징수세액 납부", law:"소득세법 제128조", color:"#FF9500" },
    { id:"wt4",  cat:"원천세", label:"원천세 납부 (4월분)",  date:`${y}-05-10`, type:"WHT", desc:"4월 급여 원천징수세액 납부", law:"소득세법 제128조", color:"#FF9500" },
    { id:"wt5",  cat:"원천세", label:"원천세 납부 (5월분)",  date:`${y}-06-10`, type:"WHT", desc:"5월 급여 원천징수세액 납부", law:"소득세법 제128조", color:"#FF9500" },
    { id:"wt6",  cat:"원천세", label:"원천세 납부 (6월분)",  date:`${y}-07-10`, type:"WHT", desc:"6월 급여 원천징수세액 납부", law:"소득세법 제128조", color:"#FF9500" },
    { id:"wt7",  cat:"원천세", label:"원천세 납부 (7월분)",  date:`${y}-08-10`, type:"WHT", desc:"7월 급여 원천징수세액 납부", law:"소득세법 제128조", color:"#FF9500" },
    { id:"wt8",  cat:"원천세", label:"원천세 납부 (8월분)",  date:`${y}-09-10`, type:"WHT", desc:"8월 급여 원천징수세액 납부", law:"소득세법 제128조", color:"#FF9500" },
    { id:"wt9",  cat:"원천세", label:"원천세 납부 (9월분)",  date:`${y}-10-10`, type:"WHT", desc:"9월 급여 원천징수세액 납부", law:"소득세법 제128조", color:"#FF9500" },
    { id:"wt10", cat:"원천세", label:"원천세 납부 (10월분)", date:`${y}-11-10`, type:"WHT", desc:"10월 급여 원천징수세액 납부", law:"소득세법 제128조", color:"#FF9500" },
    { id:"wt11", cat:"원천세", label:"원천세 납부 (11월분)", date:`${y}-12-10`, type:"WHT", desc:"11월 급여 원천징수세액 납부", law:"소득세법 제128조", color:"#FF9500" },
    { id:"wt12", cat:"원천세", label:"원천세 납부 (12월분)", date:`${y+1}-01-10`, type:"WHT", desc:"12월 급여 원천징수세액 납부", law:"소득세법 제128조", color:"#FF9500" },
    { id:"wta",  cat:"원천세", label:"연말정산 지급명세서 제출", date:`${y+1}-03-10`, type:"WHT",
      desc:"근로소득 지급명세서 제출 (2월 말 연말정산 후)", law:"소득세법 제164조", color:"#FF9500" },

    // ── 4대보험 (매월 10일·15일) ────────────────────────────
    { id:"ins1", cat:"4대보험", label:"4대보험 보험료 (1월분)",  date:`${y}-02-10`, type:"INS", desc:"국민연금·건강보험·고용보험·산재보험", law:"국민연금법 제90조 등", color:"#34C759" },
    { id:"ins2", cat:"4대보험", label:"4대보험 보험료 (2월분)",  date:`${y}-03-10`, type:"INS", desc:"국민연금·건강보험·고용보험·산재보험", law:"국민연금법 제90조 등", color:"#34C759" },
    { id:"ins3", cat:"4대보험", label:"4대보험 보험료 (3월분)",  date:`${y}-04-10`, type:"INS", desc:"국민연금·건강보험·고용보험·산재보험", law:"국민연금법 제90조 등", color:"#34C759" },
    { id:"ins4", cat:"4대보험", label:"4대보험 보험료 (4월분)",  date:`${y}-05-10`, type:"INS", desc:"국민연금·건강보험·고용보험·산재보험", law:"국민연금법 제90조 등", color:"#34C759" },
    { id:"ins5", cat:"4대보험", label:"4대보험 보험료 (5월분)",  date:`${y}-06-10`, type:"INS", desc:"국민연금·건강보험·고용보험·산재보험", law:"국민연금법 제90조 등", color:"#34C759" },
    { id:"ins6", cat:"4대보험", label:"4대보험 보험료 (6월분)",  date:`${y}-07-10`, type:"INS", desc:"국민연금·건강보험·고용보험·산재보험", law:"국민연금법 제90조 등", color:"#34C759" },
    { id:"ins7", cat:"4대보험", label:"4대보험 보험료 (7월분)",  date:`${y}-08-10`, type:"INS", desc:"국민연금·건강보험·고용보험·산재보험", law:"국민연금법 제90조 등", color:"#34C759" },
    { id:"ins8", cat:"4대보험", label:"4대보험 보험료 (8월분)",  date:`${y}-09-10`, type:"INS", desc:"국민연금·건강보험·고용보험·산재보험", law:"국민연금법 제90조 등", color:"#34C759" },
    { id:"ins9", cat:"4대보험", label:"4대보험 보험료 (9월분)",  date:`${y}-10-10`, type:"INS", desc:"국민연금·건강보험·고용보험·산재보험", law:"국민연금법 제90조 등", color:"#34C759" },
    { id:"ins10",cat:"4대보험", label:"4대보험 보험료 (10월분)", date:`${y}-11-10`, type:"INS", desc:"국민연금·건강보험·고용보험·산재보험", law:"국민연금법 제90조 등", color:"#34C759" },
    { id:"ins11",cat:"4대보험", label:"4대보험 보험료 (11월분)", date:`${y}-12-10`, type:"INS", desc:"국민연금·건강보험·고용보험·산재보험", law:"국민연금법 제90조 등", color:"#34C759" },
    { id:"ins12",cat:"4대보험", label:"4대보험 보험료 (12월분)", date:`${y+1}-01-10`, type:"INS", desc:"국민연금·건강보험·고용보험·산재보험", law:"국민연금법 제90조 등", color:"#34C759" },
    { id:"insa", cat:"4대보험", label:"건강보험 연말정산", date:`${y+1}-04-10`, type:"INS",
      desc:"전년도 보수총액 신고 후 정산", law:"국민건강보험법 제69조", color:"#34C759" },
  ];

  return items.sort((a,b)=>a.date.localeCompare(b.date));
};

// D-day 계산
const calcDday = (dateStr) => {
  const today = new Date(); today.setHours(0,0,0,0);
  const target = new Date(dateStr); target.setHours(0,0,0,0);
  return Math.round((target - today) / 86400000);
};

const MONTHS_KO = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
const TYPE_LABELS = { VAT:"부가세", CIT:"법인세", PIT:"소득세", WHT:"원천세", INS:"4대보험" };
const CAT_COLORS  = { 부가가치세:"#0071E3", 법인세:"#5E5CE6", 종합소득세:"#AF52DE", 원천세:"#FF9500", "4대보험":"#34C759" };

function TaxCalendar({ biz, bizInfo }) {
  const curYear = new Date().getFullYear();
  const [year,  setYear]   = useState(String(curYear));
  const [view,  setView]   = useState("timeline"); // timeline | calendar | list
  const [filter,setFilter] = useState("all");
  const [selMonth, setSelMonth] = useState(null); // calendar 뷰 선택 월
  const [detail, setDetail] = useState(null);     // 상세 팝업

  const bizType = bizInfo?.type || "법인";
  const schedule = buildSchedule(year, bizType);

  // 필터 적용
  const cats = ["all","부가가치세", bizType==="법인"?"법인세":"종합소득세", "원천세","4대보험"];
  const filtered = filter==="all" ? schedule : schedule.filter(s=>s.cat===filter);

  // 이번 달/다음 달 임박 항목
  const today = new Date(); today.setHours(0,0,0,0);
  const upcoming = schedule.filter(s=>{ const d=calcDday(s.date); return d>=0&&d<=30; }).slice(0,5);
  const overdue  = schedule.filter(s=>calcDday(s.date)<0&&calcDday(s.date)>=-7);

  // ── 달력 뷰용 월별 그룹
  const byMonth = {};
  filtered.forEach(s=>{
    const m = s.date.slice(0,7); // "YYYY-MM"
    if(!byMonth[m]) byMonth[m]=[];
    byMonth[m].push(s);
  });

  const DdayBadge = ({dateStr, size="sm"}) => {
    const d = calcDday(dateStr);
    const past = d < 0;
    const urgent = d <= 7 && d >= 0;
    const soon   = d <= 30 && d > 7;
    const color  = past ? T.textMuted : urgent ? T.red : soon ? T.orange : T.green;
    const bg     = past ? "rgba(0,0,0,0.05)" : urgent ? T.redLight : soon ? T.orangeLight : T.greenLight;
    const label  = past ? "완료" : d===0 ? "오늘" : `D-${d}`;
    return (
      <span style={{background:bg,color,fontSize:size==="sm"?"11px":"13px",fontWeight:"700",
        padding:size==="sm"?"2px 8px":"4px 12px",borderRadius:"20px",letterSpacing:"-0.2px",whiteSpace:"nowrap"}}>
        {label}
      </span>
    );
  };

  // ── 상세 팝업
  const DetailModal = ({item, onClose}) => (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.35)",zIndex:9000,display:"flex",alignItems:"center",justifyContent:"center",padding:"20px"}}
      onClick={onClose}>
      <Card style={{padding:"28px",maxWidth:"420px",width:"100%",boxShadow:T.shadowLg}} onClick={e=>e.stopPropagation()}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"20px"}}>
          <div>
            <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"6px"}}>
              <span style={{background:`${item.color}18`,color:item.color,fontSize:"11px",fontWeight:"700",padding:"2px 9px",borderRadius:"20px"}}>{item.cat}</span>
              <DdayBadge dateStr={item.date} size="md"/>
            </div>
            <h3 style={{color:T.text,fontSize:"17px",fontWeight:"700",margin:0,letterSpacing:"-0.5px"}}>{item.label}</h3>
          </div>
          <button onClick={onClose} style={{background:"rgba(0,0,0,0.06)",border:"none",borderRadius:"50%",width:"28px",height:"28px",cursor:"pointer",color:T.textSub,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
          </button>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:"12px"}}>
          <div style={{display:"flex",gap:"12px"}}>
            <div style={{flex:1,background:T.bg,borderRadius:T.radiusSm,padding:"12px 14px"}}>
              <p style={{color:T.textSub,fontSize:"11px",fontWeight:"600",margin:"0 0 4px",textTransform:"uppercase",letterSpacing:"0.5px"}}>납부기한</p>
              <p style={{color:T.text,fontSize:"15px",fontWeight:"700",margin:0,letterSpacing:"-0.3px"}}>
                {new Date(item.date).toLocaleDateString("ko-KR",{year:"numeric",month:"long",day:"numeric"})}
              </p>
            </div>
            <div style={{flex:1,background:T.bg,borderRadius:T.radiusSm,padding:"12px 14px"}}>
              <p style={{color:T.textSub,fontSize:"11px",fontWeight:"600",margin:"0 0 4px",textTransform:"uppercase",letterSpacing:"0.5px"}}>D-day</p>
              <p style={{color:T.text,fontSize:"15px",fontWeight:"700",margin:0}}>
                {(()=>{const d=calcDday(item.date);return d<0?"기한 경과":d===0?"오늘 마감":`${d}일 남음`;})()}
              </p>
            </div>
          </div>
          <div style={{background:T.bg,borderRadius:T.radiusSm,padding:"12px 14px"}}>
            <p style={{color:T.textSub,fontSize:"11px",fontWeight:"600",margin:"0 0 4px",textTransform:"uppercase",letterSpacing:"0.5px"}}>내용</p>
            <p style={{color:T.text,fontSize:"13px",margin:0,lineHeight:"1.6"}}>{item.desc}</p>
          </div>
          <div style={{background:"rgba(0,113,227,0.04)",border:`1px solid ${T.blueLight}`,borderRadius:T.radiusSm,padding:"12px 14px"}}>
            <p style={{color:T.textSub,fontSize:"11px",fontWeight:"600",margin:"0 0 4px",textTransform:"uppercase",letterSpacing:"0.5px"}}>근거 법령</p>
            <p style={{color:T.blue,fontSize:"13px",fontWeight:"500",margin:0}}>{item.law}</p>
          </div>
          {calcDday(item.date)>=0&&(
            <div style={{background:"rgba(255,149,0,0.06)",border:`1px solid rgba(255,149,0,0.2)`,borderRadius:T.radiusSm,padding:"12px 14px"}}>
              <p style={{color:T.orange,fontSize:"12px",fontWeight:"600",margin:"0 0 4px"}}>가산세 안내</p>
              <p style={{color:T.textSub,fontSize:"12px",margin:0,lineHeight:"1.6"}}>
                납부기한 경과 시 납부불성실 가산세(1일 0.022%) 및 신고불성실 가산세(무신고 20%, 과소신고 10%)가 부과됩니다.
                <br/>근거: 국세기본법 제47조의2·3·4
              </p>
            </div>
          )}
        </div>
      </Card>
    </div>
  );

  // ── 타임라인 뷰 (월별 그룹)
  const TimelineView = () => {
    const months = Object.keys(byMonth).sort();
    if(!months.length) return <Card style={{padding:"48px",textAlign:"center"}}><p style={{color:T.textSub,margin:0}}>해당 항목이 없습니다.</p></Card>;
    return (
      <div>
        {months.map(ym=>{
          const [y,m]=ym.split("-");
          const items=byMonth[ym];
          const isCurrentMonth = ym === `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,"0")}`;
          return (
            <div key={ym} style={{display:"flex",gap:"16px",marginBottom:"24px"}}>
              {/* 월 라벨 */}
              <div style={{width:"52px",flexShrink:0,textAlign:"center",paddingTop:"14px"}}>
                <div style={{fontSize:"11px",color:T.textMuted,fontWeight:"600"}}>{y}</div>
                <div style={{fontSize:"20px",fontWeight:"800",color:isCurrentMonth?T.blue:T.text,letterSpacing:"-0.8px",lineHeight:1.1}}>{Number(m)}월</div>
                {isCurrentMonth&&<div style={{width:"6px",height:"6px",borderRadius:"50%",background:T.blue,margin:"4px auto 0"}}/>}
              </div>
              {/* 세로 라인 */}
              <div style={{width:"2px",background:isCurrentMonth?T.blue:"rgba(0,0,0,0.07)",borderRadius:"1px",margin:"16px 0",flexShrink:0}}/>
              {/* 항목들 */}
              <div style={{flex:1,display:"flex",flexDirection:"column",gap:"6px",paddingTop:"8px"}}>
                {items.map(item=>{
                  const d=calcDday(item.date);
                  const past=d<0, urgent=d<=7&&d>=0, soon=d<=30&&d>7;
                  return (
                    <Card key={item.id}
                      onClick={()=>setDetail(item)}
                      style={{padding:"13px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:"12px",
                        borderLeft:`3px solid ${past?"rgba(0,0,0,0.1)":item.color}`,
                        background:urgent?"rgba(255,59,48,0.03)":soon?"rgba(255,149,0,0.02)":"rgba(255,255,255,0.8)",
                        opacity:past?0.55:1,
                      }}>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{display:"flex",alignItems:"center",gap:"7px",marginBottom:"3px",flexWrap:"wrap"}}>
                          <span style={{background:`${item.color}18`,color:item.color,fontSize:"10px",fontWeight:"700",padding:"1px 7px",borderRadius:"10px"}}>{item.cat}</span>
                          <span style={{color:T.text,fontSize:"13px",fontWeight:"600",letterSpacing:"-0.2px"}}>{item.label}</span>
                        </div>
                        <p style={{color:T.textMuted,fontSize:"11px",margin:0}}>
                          {new Date(item.date).toLocaleDateString("ko-KR",{month:"long",day:"numeric"})} · {item.desc}
                        </p>
                      </div>
                      <DdayBadge dateStr={item.date}/>
                    </Card>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  // ── 달력 뷰
  const CalendarView = () => {
    const [calYear, setCalYear]   = useState(Number(year));
    const [calMonth, setCalMonth] = useState(today.getMonth()); // 0-indexed

    const firstDay = new Date(calYear, calMonth, 1).getDay();
    const daysInMonth = new Date(calYear, calMonth+1, 0).getDate();
    const ym = `${calYear}-${String(calMonth+1).padStart(2,"0")}`;
    const monthItems = byMonth[ym]||[];

    // 날짜→항목 맵
    const dayMap = {};
    monthItems.forEach(item=>{
      const d = Number(item.date.split("-")[2]);
      if(!dayMap[d]) dayMap[d]=[];
      dayMap[d].push(item);
    });

    const prevMonth=()=>{ if(calMonth===0){setCalYear(y=>y-1);setCalMonth(11);}else setCalMonth(m=>m-1); };
    const nextMonth=()=>{ if(calMonth===11){setCalYear(y=>y+1);setCalMonth(0);}else setCalMonth(m=>m+1); };

    return (
      <Card style={{padding:"24px"}}>
        {/* 달력 헤더 */}
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"20px"}}>
          <button onClick={prevMonth} style={{background:"rgba(0,0,0,0.05)",border:"none",borderRadius:"8px",padding:"6px 12px",cursor:"pointer",fontSize:"14px",color:T.text}}>‹</button>
          <span style={{color:T.text,fontWeight:"700",fontSize:"17px",letterSpacing:"-0.5px"}}>{calYear}년 {MONTHS_KO[calMonth]}</span>
          <button onClick={nextMonth} style={{background:"rgba(0,0,0,0.05)",border:"none",borderRadius:"8px",padding:"6px 12px",cursor:"pointer",fontSize:"14px",color:T.text}}>›</button>
        </div>
        {/* 요일 */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",marginBottom:"6px"}}>
          {["일","월","화","수","목","금","토"].map((d,i)=>(
            <div key={d} style={{textAlign:"center",fontSize:"11px",fontWeight:"700",color:i===0?T.red:i===6?"#5E5CE6":T.textSub,padding:"4px 0"}}>{d}</div>
          ))}
        </div>
        {/* 날짜 그리드 */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:"3px"}}>
          {Array.from({length:firstDay}).map((_,i)=><div key={`e${i}`}/>)}
          {Array.from({length:daysInMonth}).map((_,i)=>{
            const day=i+1;
            const isToday=calYear===today.getFullYear()&&calMonth===today.getMonth()&&day===today.getDate();
            const items=dayMap[day]||[];
            return (
              <div key={day} style={{
                minHeight:"60px",borderRadius:"8px",padding:"5px 6px",
                background:isToday?"rgba(0,113,227,0.08)":"transparent",
                border:isToday?`1.5px solid ${T.blue}`:"1.5px solid transparent",
                cursor:items.length?"pointer":"default",
              }}
                onClick={()=>items.length===1?setDetail(items[0]):items.length>1&&setSelMonth({day,items})}
              >
                <div style={{fontSize:"12px",fontWeight:isToday?"700":"400",color:isToday?T.blue:T.text,marginBottom:"3px"}}>{day}</div>
                {items.slice(0,2).map(item=>(
                  <div key={item.id} style={{background:item.color,borderRadius:"3px",padding:"1px 4px",fontSize:"10px",color:"#fff",fontWeight:"600",marginBottom:"2px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                    {item.cat}
                  </div>
                ))}
                {items.length>2&&<div style={{fontSize:"9px",color:T.textMuted}}>+{items.length-2}개</div>}
              </div>
            );
          })}
        </div>
        {/* 선택된 날 항목 팝업 */}
        {selMonth&&(
          <div style={{marginTop:"16px",paddingTop:"16px",borderTop:`1px solid ${T.border}`}}>
            <p style={{color:T.text,fontWeight:"600",fontSize:"13px",margin:"0 0 10px"}}>{selMonth.day}일 납부 일정</p>
            {selMonth.items.map(item=>(
              <div key={item.id} onClick={()=>setDetail(item)}
                style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 14px",borderRadius:T.radiusSm,background:T.bg,marginBottom:"6px",cursor:"pointer"}}>
                <span style={{fontSize:"13px",fontWeight:"600",color:T.text}}>{item.label}</span>
                <DdayBadge dateStr={item.date}/>
              </div>
            ))}
          </div>
        )}
      </Card>
    );
  };

  // ── 리스트 뷰 (전체 표)
  const ListView = () => (
    <Card style={{overflow:"hidden"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontFamily:T.font,fontSize:"13px"}}>
        <thead><tr style={{borderBottom:`1px solid ${T.border}`,background:"rgba(0,0,0,0.02)"}}>
          {["구분","항목","납부기한","D-day","근거 법령"].map(h=>(
            <th key={h} style={{padding:"11px 16px",textAlign:"left",fontSize:"11px",fontWeight:"700",color:T.textSub,textTransform:"uppercase",letterSpacing:"0.4px"}}>{h}</th>
          ))}
        </tr></thead>
        <tbody>
          {filtered.map(item=>(
            <tr key={item.id} style={{borderBottom:`1px solid ${T.border}`,cursor:"pointer"}}
              onClick={()=>setDetail(item)}
              onMouseEnter={e=>e.currentTarget.style.background="rgba(0,0,0,0.015)"}
              onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
              <td style={{padding:"11px 16px"}}>
                <span style={{background:`${item.color}18`,color:item.color,fontSize:"11px",fontWeight:"700",padding:"2px 9px",borderRadius:"10px"}}>{item.cat}</span>
              </td>
              <td style={{padding:"11px 16px",fontWeight:"500",color:T.text}}>{item.label}</td>
              <td style={{padding:"11px 16px",color:T.textSub}}>
                {new Date(item.date).toLocaleDateString("ko-KR",{year:"numeric",month:"2-digit",day:"2-digit"})}
              </td>
              <td style={{padding:"11px 16px"}}><DdayBadge dateStr={item.date}/></td>
              <td style={{padding:"11px 16px",color:T.textSub,fontSize:"12px"}}>{item.law}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </Card>
  );

  return (
    <div>
      {/* 임박 알림 배너 */}
      {upcoming.length>0&&(
        <Card style={{padding:"16px 20px",marginBottom:"20px",background:"rgba(255,149,0,0.04)",border:`1px solid rgba(255,149,0,0.2)`}}>
          <p style={{color:T.orange,fontSize:"12px",fontWeight:"700",margin:"0 0 10px"}}>30일 이내 납부 예정</p>
          <div style={{display:"flex",flexWrap:"wrap",gap:"8px"}}>
            {upcoming.map(item=>(
              <div key={item.id} onClick={()=>setDetail(item)}
                style={{display:"flex",alignItems:"center",gap:"8px",background:"#fff",border:`1px solid ${T.border}`,borderRadius:T.radiusSm,padding:"8px 14px",cursor:"pointer",transition:"all 0.15s"}}
                onMouseEnter={e=>e.currentTarget.style.boxShadow=T.shadowMd}
                onMouseLeave={e=>e.currentTarget.style.boxShadow="none"}>
                <span style={{width:"8px",height:"8px",borderRadius:"50%",background:item.color,flexShrink:0}}/>
                <span style={{fontSize:"13px",fontWeight:"600",color:T.text}}>{item.label}</span>
                <DdayBadge dateStr={item.date}/>
              </div>
            ))}
          </div>
        </Card>
      )}

      {/* 기한 초과 경고 */}
      {overdue.length>0&&(
        <Card style={{padding:"14px 20px",marginBottom:"16px",background:T.redLight,border:`1px solid rgba(255,59,48,0.2)`}}>
          <p style={{color:T.red,fontSize:"12px",fontWeight:"700",margin:"0 0 6px"}}>기한 경과 경과 항목 — 담당자에게 즉시 문의하세요</p>
          <div style={{display:"flex",flexWrap:"wrap",gap:"6px"}}>
            {overdue.map(item=>(
              <span key={item.id} style={{fontSize:"12px",color:T.red,background:"rgba(255,59,48,0.1)",padding:"3px 10px",borderRadius:"20px",fontWeight:"600"}}>{item.label}</span>
            ))}
          </div>
        </Card>
      )}

      {/* 툴바 */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"18px",flexWrap:"wrap",gap:"10px"}}>
        {/* 연도 + 뷰 전환 */}
        <div style={{display:"flex",gap:"6px",alignItems:"center"}}>
          {[String(curYear-1),String(curYear),String(curYear+1)].map(y=>(
            <button key={y} onClick={()=>setYear(y)}
              style={{padding:"6px 14px",borderRadius:"8px",border:"none",cursor:"pointer",fontSize:"13px",fontFamily:T.font,fontWeight:"600",
                background:year===y?T.blue:"rgba(0,0,0,0.05)",color:year===y?"#fff":T.textSub,transition:"all 0.15s"}}>
              {y}
            </button>
          ))}
          <span style={{color:T.border,margin:"0 4px"}}>|</span>
          {[{id:"timeline",label:"타임라인"},{id:"calendar",label:"달력"},{id:"list",label:"목록"}].map(v=>(
            <button key={v.id} onClick={()=>setView(v.id)}
              style={{padding:"6px 13px",borderRadius:"8px",border:"none",cursor:"pointer",fontSize:"12px",fontFamily:T.font,fontWeight:"600",
                background:view===v.id?"rgba(0,113,227,0.12)":"rgba(0,0,0,0.04)",color:view===v.id?T.blue:T.textSub}}>
              {v.label}
            </button>
          ))}
        </div>
        {/* 카테고리 필터 */}
        <div style={{display:"flex",gap:"5px",flexWrap:"wrap"}}>
          {cats.map(c=>(
            <button key={c} onClick={()=>setFilter(c)}
              style={{padding:"5px 12px",borderRadius:"20px",border:"none",cursor:"pointer",fontSize:"12px",fontFamily:T.font,fontWeight:"600",
                background:filter===c?(c==="all"?T.blue:CAT_COLORS[c]||T.blue):"rgba(0,0,0,0.05)",
                color:filter===c?"#fff":T.textSub,transition:"all 0.15s"}}>
              {c==="all"?"전체":c}
            </button>
          ))}
        </div>
      </div>

      {/* 뷰 */}
      {view==="timeline" && <TimelineView/>}
      {view==="calendar" && <CalendarView/>}
      {view==="list"     && <ListView/>}

      {/* 상세 팝업 */}
      {detail && <DetailModal item={detail} onClose={()=>setDetail(null)}/>}

      {/* 면책 고지 */}
      <p style={{color:T.textMuted,fontSize:"11px",margin:"20px 0 0",lineHeight:"1.7"}}>
        ※ 위 일정은 일반적인 기준일이며, 납기일이 토·일·공휴일인 경우 다음 영업일로 연장됩니다 (국세기본법 제5조).<br/>
        ※ 반기납부 특례 사업자(원천세), 간이과세자 등은 납부 일정이 다를 수 있습니다. 정확한 일정은 담당자에게 확인하세요.
      </p>
    </div>
  );
}

/* ─────────────────────────────────────────
   관리자 업로드 패널
───────────────────────────────────────── */
function UploadPanel({ businesses, onRefresh, showToast }) {
  const [bizNo,setBizNo]=useState("");
  const [year,setYear]=useState("2024");
  const [docType,setDocType]=useState("재무상태표");
  const [parsed,setParsed]=useState(null);
  const [preview,setPreview]=useState(false);

  const handleParsed=(data)=>{ setParsed(data); setPreview(true); };
  const handleSave=()=>{
    if(!bizNo||!parsed){showToast("사업자와 파일을 선택해주세요.");return;}
    db.saveDoc(bizNo,year,docType,parsed)
      .then(()=>{ showToast("저장 완료 — 고객이 즉시 조회 가능합니다."); setParsed(null); setPreview(false); if(onRefresh) onRefresh(); })
      .catch(()=>showToast("저장 중 오류가 발생했습니다."));
  };

  const selStyle={padding:"10px 14px",borderRadius:T.radiusSm,border:`1.5px solid ${T.border}`,background:"rgba(255,255,255,0.9)",color:T.text,fontSize:"14px",fontFamily:T.font,outline:"none",width:"100%",boxSizing:"border-box"};

  return (
    <div>
      <div style={{marginBottom:"24px"}}>
        <h2 style={{color:T.text,fontSize:"20px",fontWeight:"700",margin:"0 0 6px",letterSpacing:"-0.7px"}}>Excel 업로드</h2>
        <p style={{color:T.textSub,fontSize:"14px",margin:0}}>위하고T에서 엑셀 변환한 파일을 업로드하면 자동으로 파싱됩니다.</p>
      </div>

      {/* 안내 박스 */}
      <Card style={{padding:"16px 20px",marginBottom:"20px",background:"rgba(0,113,227,0.04)",border:`1px solid ${T.blueLight}`}}>
        <p style={{color:T.blue,fontSize:"12px",fontWeight:"700",margin:"0 0 8px"}}>위하고T 엑셀 출력 방법</p>
        <div style={{color:T.textSub,fontSize:"12px",lineHeight:"1.9"}}>
          <p style={{margin:0}}>① 전체메뉴 → 결산재무제표 → 해당 항목 클릭</p>
          <p style={{margin:0}}>② 과목별 탭 선택 → 결산월 조회</p>
          <p style={{margin:0}}>③ 마우스 우클릭 → 데이터 변환 → <strong>엑셀변환</strong> 선택</p>
        </div>
      </Card>

      <Card style={{padding:"28px",marginBottom:"20px"}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"14px",marginBottom:"20px"}}>
          <div>
            <Label>사업자</Label>
            <select value={bizNo} onChange={e=>setBizNo(e.target.value)} style={selStyle}>
              <option value="">선택</option>
              {Object.entries(businesses).map(([no,info])=><option key={no} value={no}>{info.name}</option>)}
            </select>
          </div>
          <div>
            <Label>사업연도</Label>
            <select value={year} onChange={e=>setYear(e.target.value)} style={selStyle}>
              {["2025","2024","2023","2022","2021"].map(y=><option key={y} value={y}>{y}년</option>)}
            </select>
          </div>
          <div>
            <Label>자료 종류</Label>
            <select value={docType} onChange={e=>{setDocType(e.target.value);setParsed(null);setPreview(false);}} style={selStyle}>
              {["재무상태표","손익계산서","부가세신고","법인세신고"].map(t=><option key={t} value={t}>{t}</option>)}
            </select>
          </div>
        </div>

        <ExcelDropzone onParsed={handleParsed} docType={docType}/>

        {parsed && (
          <div style={{marginTop:"16px",display:"flex",gap:"10px",alignItems:"center"}}>
            <Btn onClick={()=>setPreview(!preview)} variant="secondary" size="sm">{preview?"미리보기 닫기":"미리보기 열기"}</Btn>
            <Btn onClick={handleSave} size="sm">저장하기</Btn>
            <span style={{color:T.textSub,fontSize:"13px"}}>파싱 완료 — 저장 후 고객이 즉시 조회 가능합니다</span>
          </div>
        )}
      </Card>

      {/* 미리보기 */}
      {preview && parsed && (
        <div>
          <p style={{color:T.text,fontWeight:"600",fontSize:"14px",margin:"0 0 12px",letterSpacing:"-0.3px"}}>미리보기</p>
          {docType==="재무상태표"&&<BSView data={parsed}/>}
          {docType==="손익계산서"&&<ISView data={parsed}/>}
          {docType==="부가세신고"&&<VATView data={parsed}/>}
          {docType==="법인세신고"&&<TaxView data={parsed}/>}
        </div>
      )}
    </div>
  );
}

/* ─────────────────────────────────────────
   고객 대시보드
───────────────────────────────────────── */
const MENUS = [
  {id:"재무상태표", label:"재무상태표",    icon:"balanceSheet"},
  {id:"손익계산서", label:"손익계산서",    icon:"income"},
  {id:"부가세신고", label:"부가세 신고서", icon:"vat"},
  {id:"법인세신고", label:"법인세·소득세", icon:"tax"},
  {id:"세금일정",   label:"세금 납부 일정",icon:"calendar"},
];

function CustomerDash({ user, onLogout }) {
  const isMobile = useIsMobile();
  const [biz,setBiz]=useState(user.businesses?.[0]||"");
  const [year,setYear]=useState("");           // ← 빈 문자열로 시작, getYears 후 자동 설정
  const [menu,setMenu]=useState("재무상태표");
  const [bizInfo,setBizInfo]=useState(null);
  const [allBizInfo,setAllBizInfo]=useState({}); // ← 전체 사업자 정보 캐시 (모바일 BottomSheet용)
  const [years,setYears]=useState([]);
  const [sidebar,setSidebar]=useState(true);
  const [sheetOpen,setSheetOpen]=useState(false);
  const [doc,setDoc]=useState(null);
  const [prevDoc,setPrevDoc]=useState(null);
  const [loading,setLoading]=useState(false);

  // 최초 마운트 시 전체 사업자 정보 로드 (모바일 사업자 목록 표시용)
  useEffect(()=>{
    if(!user.businesses?.length) return;
    Promise.all(user.businesses.map(b=>db.getBiz(b).catch(()=>null)))
      .then(results=>{
        const m={};
        user.businesses.forEach((b,i)=>{ if(results[i]) m[b]=results[i]; });
        setAllBizInfo(m);
      });
  },[]);

  // 사업자 변경 시 기본 정보 로드
  useEffect(()=>{
    if(!biz) return;
    db.getBiz(biz).then(info=>setBizInfo(info)).catch(()=>{});
    db.getYears(biz).then(ys=>{
      setYears(ys);
      // 현재 연도가 목록에 없으면 첫 번째 연도로 자동 설정
      if(ys.length) setYear(prev => ys.includes(prev) ? prev : ys[0]);
    }).catch(()=>{});
  },[biz]);

  // 메뉴/연도/사업자 변경 시 데이터 로드
  useEffect(()=>{
    if(!biz || !year || menu==="세금일정") { setDoc(null); setPrevDoc(null); return; }
    setLoading(true);
    const prevYear = year ? String(Number(year)-1) : null;
    Promise.all([
      db.getDoc(biz, year, menu),
      prevYear ? db.getDoc(biz, prevYear, menu) : Promise.resolve(null),
    ]).then(([d, pd])=>{ setDoc(d); setPrevDoc(pd); setLoading(false); })
      .catch(()=>{ setDoc(null); setPrevDoc(null); setLoading(false); });
  },[biz, year, menu]);

  const prevYear = year ? String(Number(year)-1) : null;

  // ── 모바일 하단 탭바 (아이콘 5개)
  const BottomTab = () => (
    <nav style={{
      position:"fixed",bottom:0,left:0,right:0,zIndex:200,
      background:T.bgDeep,
      borderTop:`1px solid ${T.borderNav}`,
      display:"flex",height:"60px",paddingBottom:"env(safe-area-inset-bottom)",
    }} className="no-print">
      {MENUS.map(m=>(
        <button key={m.id} onClick={()=>{setMenu(m.id);setSheetOpen(false);}} style={{
          flex:1,display:"flex",flexDirection:"column",alignItems:"center",
          justifyContent:"center",gap:"3px",border:"none",
          background:"transparent",cursor:"pointer",
          borderTop: menu===m.id?`2px solid ${T.gold}`:"2px solid transparent",
          paddingTop:"4px",
        }}>
          <Icon name={m.icon} size={18} color={menu===m.id?T.gold:"rgba(255,255,255,0.4)"}/>
          <span style={{fontSize:"9px",fontWeight:menu===m.id?"700":"400",
            color:menu===m.id?T.gold:"rgba(255,255,255,0.4)",letterSpacing:"-0.1px"}}>
            {m.label.replace(" 신고서","").replace("·소득세","")}
          </span>
        </button>
      ))}
      <button onClick={()=>setSheetOpen(true)} style={{
        flex:1,display:"flex",flexDirection:"column",alignItems:"center",
        justifyContent:"center",gap:"3px",border:"none",background:"transparent",cursor:"pointer",
        borderTop:"2px solid transparent",paddingTop:"4px",
      }}>
        <Icon name="settings" size={18} color="rgba(255,255,255,0.4)"/>
        <span style={{fontSize:"9px",fontWeight:"400",color:"rgba(255,255,255,0.4)"}}>설정</span>
      </button>
    </nav>
  );

  // ── 모바일 하단 슬라이드 시트 (연도·사업자 선택)
  const BottomSheet = () => (
    <>
      <div onClick={()=>setSheetOpen(false)}
        style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.3)",zIndex:300}}/>
      <div style={{
        position:"fixed",bottom:0,left:0,right:0,zIndex:400,
        background:"#fff",borderRadius:"20px 20px 0 0",
        padding:"20px 20px calc(20px + env(safe-area-inset-bottom))",
        boxShadow:"0 -8px 40px rgba(0,0,0,0.15)",
        animation:"slideUp 0.25s ease",
      }}>
        <style>{`@keyframes slideUp{from{transform:translateY(100%)}to{transform:translateY(0)}}`}</style>
        {/* 핸들 */}
        <div style={{width:"36px",height:"4px",borderRadius:"2px",background:"rgba(0,0,0,0.15)",margin:"0 auto 20px"}}/>

        {/* 사업자 선택 */}
        {user.businesses?.length>1&&(
          <div style={{marginBottom:"20px"}}>
            <p style={{color:T.textSub,fontSize:"12px",fontWeight:"700",textTransform:"uppercase",letterSpacing:"0.6px",margin:"0 0 10px"}}>사업자 선택</p>
            {user.businesses.map(b=>{
              const info=allBizInfo[b];
              return (
                <button key={b} onClick={()=>{setBiz(b);setSheetOpen(false);}} style={{
                  width:"100%",textAlign:"left",padding:"12px 14px",borderRadius:T.radiusSm,
                  border:`1.5px solid ${biz===b?T.blue:T.border}`,marginBottom:"6px",
                  background:biz===b?T.blueLight:"#fff",cursor:"pointer",fontFamily:T.font,
                  display:"flex",alignItems:"center",justifyContent:"space-between",
                }}>
                  <div>
                    <p style={{color:T.text,fontSize:"14px",fontWeight:"600",margin:"0 0 2px"}}>{info?.name||b}</p>
                    <p style={{color:T.textSub,fontSize:"12px",margin:0}}>{b} · {info?.type}</p>
                  </div>
                  {biz===b&&<Icon name="check" size={16} color={T.blue} strokeWidth={2.5}/>}
                </button>
              );
            })}
          </div>
        )}

        {/* 연도 선택 */}
        <div>
          <p style={{color:T.textSub,fontSize:"12px",fontWeight:"700",textTransform:"uppercase",letterSpacing:"0.6px",margin:"0 0 10px"}}>사업연도</p>
          <div style={{display:"flex",gap:"8px",flexWrap:"wrap"}}>
            {years.map(y=>(
              <button key={y} onClick={()=>{setYear(y);setSheetOpen(false);}} style={{
                padding:"10px 20px",borderRadius:T.radiusSm,border:"none",cursor:"pointer",
                fontSize:"15px",fontFamily:T.font,fontWeight:"600",
                background:year===y?T.blue:"rgba(0,0,0,0.06)",
                color:year===y?"#fff":T.textSub,
              }}>{y}년</button>
            ))}
          </div>
        </div>
      </div>
    </>
  );

  return (
    <div style={{minHeight:"100vh",background:T.bg,fontFamily:T.font,display:"flex",flexDirection:"column"}}>
      {/* ── 헤더 ── */}
      <header style={{
        background:T.bgDeep,
        borderBottom:`1px solid ${T.borderNav}`,
        padding:`0 ${isMobile?"14px":"20px"}`,height:"54px",
        display:"flex",alignItems:"center",justifyContent:"space-between",
        position:"sticky",top:0,zIndex:100,
      }}>
        <div style={{display:"flex",alignItems:"center",gap:"12px"}}>
          {!isMobile&&(
            <button onClick={()=>setSidebar(!sidebar)} style={{background:"none",border:"none",color:"rgba(255,255,255,0.5)",cursor:"pointer",padding:"4px 6px",borderRadius:"6px",display:"flex",alignItems:"center"}}>
              <Icon name="menu" size={18} color="rgba(255,255,255,0.6)"/>
            </button>
          )}
          <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
            <div style={{width:"28px",height:"28px",borderRadius:"6px",background:`linear-gradient(135deg,${T.gold},#A8843A)`,display:"flex",alignItems:"center",justifyContent:"center"}}>
              <Icon name="logo" size={14} color="#fff" strokeWidth={2.5}/>
            </div>
            <span style={{color:"#fff",fontWeight:"800",fontSize:isMobile?"13px":"15px",letterSpacing:"-0.4px"}}>
              {isMobile?(bizInfo?.name||"고객 전용 재무포털"):"고객 전용 재무정보 조회 시스템"}
            </span>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
          {!isMobile&&user.businesses?.length>1&&(
            <select value={biz} onChange={e=>{setBiz(e.target.value);}} style={{background:"rgba(255,255,255,0.1)",border:"1px solid rgba(255,255,255,0.15)",color:"#fff",padding:"6px 10px",borderRadius:"7px",fontSize:"13px",cursor:"pointer",fontFamily:T.font,fontWeight:"500"}}>
              {user.businesses.map(b=>{ return <option key={b} value={b} style={{background:T.bgDeep}}>{b}</option>; })}
            </select>
          )}
          {!isMobile&&<span style={{color:"rgba(255,255,255,0.6)",fontSize:"13px",fontWeight:"500"}}>{user.name}</span>}
          <button onClick={onLogout} style={{background:"rgba(255,255,255,0.1)",border:"1px solid rgba(255,255,255,0.15)",color:"rgba(255,255,255,0.8)",padding:"5px 13px",borderRadius:"7px",fontSize:"12px",fontWeight:"600",cursor:"pointer",fontFamily:T.font,display:"flex",alignItems:"center",gap:"5px"}}>
            <Icon name="logout" size={13} color="rgba(255,255,255,0.8)"/>
            {!isMobile&&"로그아웃"}
          </button>
        </div>
      </header>

      <div style={{display:"flex",flex:1}}>
        {/* ── PC 사이드바 ── */}
        {!isMobile&&sidebar&&(
          <aside style={{width:"210px",background:T.bgDeepAlt,borderRight:`1px solid ${T.borderNav}`,padding:"20px 14px",flexShrink:0}}>
            {/* 사업자 카드 */}
            <div style={{background:"rgba(255,255,255,0.07)",border:`1px solid ${T.borderNav}`,borderRadius:T.radiusSm,padding:"12px 14px",marginBottom:"20px"}}>
              <p style={{color:T.gold,fontSize:"9px",fontWeight:"700",margin:"0 0 5px",textTransform:"uppercase",letterSpacing:"1px"}}>현재 사업자</p>
              <p style={{color:"#fff",fontSize:"13px",fontWeight:"700",margin:"0 0 3px",letterSpacing:"-0.3px"}}>{bizInfo?.name||"-"}</p>
              <p style={{color:"rgba(255,255,255,0.45)",fontSize:"11px",margin:"0 0 1px"}}>{biz}</p>
              <p style={{color:"rgba(255,255,255,0.35)",fontSize:"10px",margin:0}}>{bizInfo?.type}</p>
            </div>
            <p style={{color:"rgba(255,255,255,0.3)",fontSize:"9px",fontWeight:"700",textTransform:"uppercase",letterSpacing:"1px",margin:"0 0 8px",paddingLeft:"4px"}}>사업연도</p>
            <div style={{display:"flex",flexWrap:"wrap",gap:"5px",marginBottom:"20px"}}>
              {years.map(y=>(
                <button key={y} onClick={()=>setYear(y)} style={{
                  padding:"5px 12px",borderRadius:"7px",border:"none",cursor:"pointer",
                  fontSize:"12px",fontFamily:T.font,fontWeight:"600",
                  background:year===y?T.gold:"rgba(255,255,255,0.08)",
                  color:year===y?T.bgDeep:"rgba(255,255,255,0.88)",
                  transition:"all 0.15s",
                }}>{y}</button>
              ))}
            </div>
            <p style={{color:"rgba(255,255,255,0.3)",fontSize:"9px",fontWeight:"700",textTransform:"uppercase",letterSpacing:"1px",margin:"0 0 6px",paddingLeft:"4px"}}>조회 메뉴</p>
            {MENUS.map(m=>(
              <button key={m.id} onClick={()=>setMenu(m.id)} style={{
                width:"100%",textAlign:"left",padding:"9px 12px",borderRadius:T.radiusSm,
                border:"none",cursor:"pointer",fontSize:"13px",fontFamily:T.font,
                fontWeight:menu===m.id?"700":"600",marginBottom:"2px",
                display:"flex",alignItems:"center",gap:"9px",
                background:menu===m.id?T.gold:"transparent",
                color:menu===m.id?T.bgDeep:"#FFFFFF",
                letterSpacing:"-0.2px",transition:"all 0.15s",
                borderLeft:menu===m.id?"none":"2px solid transparent",
              }}
                onMouseEnter={e=>{if(menu!==m.id){e.currentTarget.style.background="rgba(255,255,255,0.08)";e.currentTarget.style.color="#FFFFFF";}}}
                onMouseLeave={e=>{if(menu!==m.id){e.currentTarget.style.background="transparent";e.currentTarget.style.color="#FFFFFF";}}}
              >
                <Icon name={m.icon} size={15} color={menu===m.id?T.bgDeep:"#FFFFFF"}/>
                <span>{m.label}</span>
              </button>
            ))}
          </aside>
        )}

        {/* ── 메인 컨텐츠 ── */}
        <main style={{
          flex:1,
          padding: isMobile ? "16px 14px 80px" : "28px",
          overflow:"auto",minWidth:0,
        }}>
          {/* 인쇄용 CSS + 모바일 반응형 CSS */}
          <style>{`
            @media print {
              @page { size: A4; margin: 15mm 12mm; }
              body * { visibility: hidden !important; }
              #print-area, #print-area * { visibility: visible !important; }
              #print-area {
                position: fixed !important; top: 0 !important; left: 0 !important;
                width: 100% !important; padding: 0 !important;
                background: #fff !important;
                font-family: -apple-system, 'Helvetica Neue', sans-serif !important;
              }
              #print-header { display: flex !important; }
              .no-print { display: none !important; }
              table { border-collapse: collapse !important; width: 100% !important; margin-bottom: 16px !important; }
              th { background: #f0f4ff !important; color: #1D1D1F !important; font-size: 11px !important; font-weight: 700 !important; padding: 8px 12px !important; border: 1px solid #dde3f0 !important; text-align: left !important; }
              td { font-size: 12px !important; padding: 8px 12px !important; border: 1px solid #e8e8ed !important; color: #1D1D1F !important; }
              tr:nth-child(even) td { background: #fafafa !important; }
              tr { page-break-inside: avoid !important; }
              .kpi-grid { display: flex !important; flex-wrap: wrap !important; gap: 8px !important; margin-bottom: 16px !important; }
              .chart-area { display: none !important; }
              .two-col { display: block !important; }
              .two-col > * { width: 100% !important; margin-bottom: 16px !important; }
            }
            #print-header { display: none; }
            /* 모바일 테이블 가로 스크롤 */
            .mob-scroll { overflow-x: auto; -webkit-overflow-scrolling: touch; }
            .mob-scroll table { min-width: 540px; }
          `}</style>

          {/* 페이지 헤더 */}
          <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:isMobile?"16px":"22px"}}>
            <div>
              <h2 style={{color:T.text,fontSize:isMobile?"17px":"20px",fontWeight:"700",margin:"0 0 4px",letterSpacing:"-0.7px"}}>
                {MENUS.find(m=>m.id===menu)?.icon} {MENUS.find(m=>m.id===menu)?.label}
              </h2>
              <p style={{color:T.textSub,fontSize:"12px",margin:0}}>
                {bizInfo?.name||"—"}{menu!=="세금일정"&&` · ${year}년`}
              </p>
            </div>
            {doc && menu!=="세금일정" && (
              <button
                className="no-print"
                onClick={()=>{
                  const prev=document.title;
                  document.title=`${bizInfo?.name||""}_${year}_${MENUS.find(m=>m.id===menu)?.label}`;
                  window.print();
                  document.title=prev;
                }}
                style={{
                  display:"flex",alignItems:"center",gap:"6px",
                  padding: isMobile?"8px 12px":"9px 18px",
                  borderRadius:T.radiusSm,border:"none",
                  background:T.blue,color:"#fff",
                  fontSize: isMobile?"12px":"13px",
                  fontWeight:"600",fontFamily:T.font,cursor:"pointer",
                  boxShadow:"0 2px 8px rgba(0,113,227,0.3)",
                  letterSpacing:"-0.2px",flexShrink:0,
                }}
              >
                <Icon name="download" size={14} color="#fff"/>{!isMobile&&" PDF 출력"}
              </button>
            )}
          </div>

          {/* 인쇄 영역 */}
          <div id="print-area">
            <div id="print-header" style={{marginBottom:"20px",paddingBottom:"12px",borderBottom:"2px solid #1D1D1F"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end"}}>
                <div>
                  <div style={{fontSize:"18px",fontWeight:"700",marginBottom:"4px"}}>{bizInfo?.name}</div>
                  <div style={{fontSize:"13px",color:"#6E6E73"}}>사업자번호: {biz} · {bizInfo?.type}</div>
                </div>
                <div style={{textAlign:"right"}}>
                  <div style={{fontSize:"16px",fontWeight:"700"}}>{MENUS.find(m=>m.id===menu)?.label}</div>
                  <div style={{fontSize:"12px",color:"#6E6E73",marginTop:"2px"}}>출력일: {new Date().toLocaleDateString("ko-KR")} · {year}년</div>
                </div>
              </div>
            </div>

            {menu==="세금일정" ? (
              <TaxCalendar biz={biz} bizInfo={bizInfo}/>
            ) : loading ? (
              <div style={{...CARD_STYLE,padding:"60px",textAlign:"center"}}>
                <div style={{width:"32px",height:"32px",border:`3px solid ${T.border}`,borderTop:`3px solid ${T.gold}`,borderRadius:"50%",margin:"0 auto 16px",animation:"spin 0.8s linear infinite"}}/>
                <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
                <p style={{color:T.textMuted,fontSize:"14px",margin:0}}>데이터를 불러오는 중...</p>
              </div>
            ) : !doc ? <Empty/> : (
              <>
                {menu==="재무상태표"&&<BSView data={doc} prevData={prevDoc} prevYear={prevYear} isMobile={isMobile}/>}
                {menu==="손익계산서"&&<ISView data={doc} prevData={prevDoc} prevYear={prevYear} isMobile={isMobile}/>}
                {menu==="부가세신고"&&<VATView data={doc} prevData={prevDoc} prevYear={prevYear} isMobile={isMobile}/>}
                {menu==="법인세신고"&&<TaxView data={doc} prevData={prevDoc} prevYear={prevYear} isMobile={isMobile}/>}
              </>
            )}
          </div>
        </main>
      </div>

      {/* ── 모바일 하단 탭바 ── */}
      {isMobile&&<BottomTab/>}
      {isMobile&&sheetOpen&&<BottomSheet/>}
    </div>
  );
}

/* ─────────────────────────────────────────
   관리자 대시보드
───────────────────────────────────────── */
function AdminDash({ user, onLogout }) {
  const [tab,setTab]=useState("upload");
  const [users,setUsers]=useState([]);
  const [businesses,setBusinesses]=useState({});
  const [toast,setToast]=useState("");
  const refresh=useCallback(()=>{
    db.getUsers().then(u=>setUsers(u)).catch(()=>{});
    db.getAllBizNos().then(b=>setBusinesses(b)).catch(()=>{});
  },[]);
  useEffect(()=>{ refresh(); const id=setInterval(refresh,30000); return()=>clearInterval(id); },[refresh]);
  const showToast=(m)=>{ setToast(m); setTimeout(()=>setToast(""),2800); };
  const pending=users.filter(u=>u.status==="pending").length;
  const TABS=[{id:"upload",label:"업로드",icon:"upload"},{id:"members",label:"회원 관리",icon:"users",badge:pending},{id:"businesses",label:"사업자",icon:"building"},{id:"edit",label:"회원 정보 수정",icon:"settings"}];

  return (
    <div style={{minHeight:"100vh",background:T.bg,fontFamily:T.font}}>
      {toast&&<div style={{position:"fixed",top:"20px",right:"20px",zIndex:9999,background:"rgba(255,255,255,0.96)",border:`1px solid ${T.border}`,borderRadius:T.radius,padding:"12px 20px",color:T.text,fontSize:"14px",fontWeight:"500",boxShadow:T.shadowMd,backdropFilter:"blur(20px)"}}>{toast}</div>}
      <header style={{background:T.bgDeep,borderBottom:`1px solid ${T.borderNav}`,padding:"0 24px",height:"52px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div style={{display:"flex",alignItems:"center",gap:"10px"}}>
          <span style={{color:"#fff",fontWeight:"700",fontSize:"15px",letterSpacing:"-0.4px"}}>고객 전용 재무정보 조회 시스템</span>
          <Pill label="관리자"/>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:"10px"}}>
          <span style={{color:T.textSub,fontSize:"13px"}}>{user.name}</span>
          <Btn onClick={onLogout} variant="secondary" size="sm">로그아웃</Btn>
        </div>
      </header>
      <div style={{display:"flex"}}>
        <aside style={{width:"190px",background:T.bgDeepAlt,borderRight:`1px solid ${T.border}`,padding:"20px 12px",minHeight:"calc(100vh - 52px)"}}>
          <div style={{marginBottom:"20px"}}>
            {[{l:"전체",v:users.length},{l:"대기",v:pending,c:T.orange},{l:"승인",v:users.filter(u=>u.status==="approved").length,c:T.green}].map(s=>(
              <div key={s.l} style={{display:"flex",justifyContent:"space-between",padding:"7px 10px",borderRadius:"8px",marginBottom:"3px"}}>
                <span style={{color:T.textSub,fontSize:"12px"}}>{s.l}</span>
                <span style={{color:s.c||T.text,fontWeight:"700",fontSize:"13px"}}>{s.v}</span>
              </div>
            ))}
          </div>
          {TABS.map(t=>(
            <button key={t.id} onClick={()=>setTab(t.id)} style={{width:"100%",textAlign:"left",padding:"9px 12px",borderRadius:T.radiusSm,border:"none",cursor:"pointer",fontSize:"13px",fontFamily:T.font,fontWeight:tab===t.id?"700":"600",marginBottom:"2px",display:"flex",alignItems:"center",gap:"8px",justifyContent:"space-between",background:tab===t.id?"rgba(255,255,255,0.12)":"transparent",color:tab===t.id?"#FFFFFF":"rgba(255,255,255,0.85)"}}>
              <span style={{display:"flex",alignItems:"center",gap:"8px"}}><Icon name={t.icon} size={14} color={tab===t.id?"#FFFFFF":"rgba(255,255,255,0.7)"}/><span>{t.label}</span></span>
              {t.badge>0&&<span style={{background:T.red,color:"#fff",fontSize:"10px",fontWeight:"700",padding:"1px 6px",borderRadius:"10px"}}>{t.badge}</span>}
            </button>
          ))}
        </aside>
        <main style={{flex:1,padding:"28px",overflow:"auto"}}>
          {tab==="upload"&&<UploadPanel businesses={businesses} onRefresh={refresh} showToast={showToast}/>}
          {tab==="members"&&<MembersPanel users={users} businesses={businesses} onRefresh={refresh} showToast={showToast}/>}
          {tab==="businesses"&&<BizPanel businesses={businesses} onRefresh={refresh} showToast={showToast}/>}
          {tab==="edit"&&<EditMembersPanel users={users} businesses={businesses} onRefresh={refresh} showToast={showToast}/>}
        </main>
      </div>
    </div>
  );
}

function MembersPanel({ users, businesses, onRefresh, showToast }) {
  const [filter,setFilter]=useState("all");
  const [editId,setEditId]=useState(null);
  const [newBiz,setNewBiz]=useState("");
  const filtered=filter==="all"?users:users.filter(u=>u.status===filter);
  const badge=(s)=>{ const m={pending:["대기",T.orange,T.orangeLight],approved:["승인",T.green,T.greenLight],rejected:["거절",T.red,T.redLight]}; const[l,c,bg]=m[s]||[]; return <Pill label={l} color={c} bg={bg}/>; };
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"20px"}}>
        <h2 style={{color:T.text,fontSize:"20px",fontWeight:"700",margin:0,letterSpacing:"-0.7px"}}>회원 관리</h2>
        <div style={{display:"flex",gap:"6px"}}>
          {[["all","전체"],["pending","대기"],["approved","승인"],["rejected","거절"]].map(([v,l])=>(
            <Btn key={v} onClick={()=>setFilter(v)} variant={filter===v?"primary":"secondary"} size="sm">{l}</Btn>
          ))}
        </div>
      </div>
      {filtered.length===0&&<Card style={{padding:"48px",textAlign:"center"}}><p style={{color:T.textSub,margin:0}}>해당하는 회원이 없습니다.</p></Card>}
      {filtered.map(u=>(
        <Card key={u.id} style={{padding:"20px",marginBottom:"10px"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div>
              <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"4px"}}>
                <span style={{color:T.text,fontWeight:"700",fontSize:"15px",letterSpacing:"-0.4px"}}>{u.name}</span>{badge(u.status)}
              </div>
              <p style={{color:T.textSub,fontSize:"13px",margin:"0 0 2px"}}>{u.email}{u.phone?` · ${u.phone}`:""}</p>
              {u.memo&&<p style={{color:T.textSub,fontSize:"12px",margin:"0 0 2px"}}>  {u.memo}</p>}
              {u.registeredAt&&<p style={{color:T.textMuted,fontSize:"11px",margin:0}}>신청일: {new Date(u.registeredAt).toLocaleDateString("ko-KR")}</p>}
            </div>
            {u.status==="pending"&&(
              <div style={{display:"flex",gap:"6px"}}>
                <Btn onClick={()=>{db.setStatus(u.id,"approved").then(()=>{onRefresh();showToast("승인 처리되었습니다.");}).catch(()=>showToast("오류가 발생했습니다."));}} variant="secondary" size="sm" style={{color:T.green}}>승인</Btn>
                <Btn onClick={()=>{db.setStatus(u.id,"rejected").then(()=>{onRefresh();showToast("거절되었습니다.");}).catch(()=>showToast("오류가 발생했습니다."));}} variant="secondary" size="sm" style={{color:T.red}}>거절</Btn>
              </div>
            )}
            {/* 삭제 버튼 — 항상 표시 */}
            <Btn onClick={()=>{
              if(!window.confirm(`"${u.name}" 회원을 삭제하시겠습니까?\n삭제 시 복구가 불가능합니다.`)) return;
              db.deleteUser(u.id).then(()=>{onRefresh();showToast("삭제되었습니다.");}).catch(()=>showToast("삭제 중 오류가 발생했습니다."));
            }} variant="secondary" size="sm" style={{color:T.red,border:`1px solid ${T.red}`,marginLeft:"4px"}}>삭제</Btn>
          </div>
          <div style={{marginTop:"14px",paddingTop:"14px",borderTop:`1px solid ${T.border}`}}>
            <p style={{color:T.textMuted,fontSize:"10px",fontWeight:"700",textTransform:"uppercase",letterSpacing:"0.6px",margin:"0 0 8px"}}>연결된 사업자</p>
            <div style={{display:"flex",flexWrap:"wrap",gap:"6px",marginBottom:"8px"}}>
              {(u.businesses||[]).map(biz=>(
                <span key={biz} style={{display:"inline-flex",alignItems:"center",gap:"5px",background:T.blueLight,color:T.blue,fontSize:"12px",fontWeight:"500",padding:"3px 10px",borderRadius:"20px"}}>
                  {businesses[biz]?.name||biz}
                  <button onClick={()=>{db.setBizList(u.id,u.businesses.filter(b=>b!==biz)).then(()=>{onRefresh();}).catch(()=>{});}} style={{background:"none",border:"none",color:T.red,cursor:"pointer",padding:"0 0 0 2px",display:"inline-flex",alignItems:"center"}}>
                    <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
                  </button>
                </span>
              ))}
              {!(u.businesses?.length)&&<span style={{color:T.textMuted,fontSize:"12px"}}>없음</span>}
            </div>
            {editId===u.id?(
              <div style={{display:"flex",gap:"6px"}}>
                <Input value={newBiz} onChange={setNewBiz} placeholder="사업자번호 입력" style={{flex:1}}/>
                <Btn onClick={()=>{if(!newBiz.trim())return;db.setBizList(u.id,[...(u.businesses||[]),newBiz.trim()]).then(()=>{setNewBiz("");setEditId(null);onRefresh();showToast("추가되었습니다.");}).catch(()=>showToast("오류가 발생했습니다."));}} size="sm">추가</Btn>
                <Btn onClick={()=>setEditId(null)} variant="secondary" size="sm">취소</Btn>
              </div>
            ):(
              <button onClick={()=>setEditId(u.id)} style={{background:"none",border:`1px dashed ${T.borderStrong}`,color:T.textSub,padding:"4px 12px",borderRadius:"8px",cursor:"pointer",fontSize:"12px",fontFamily:T.font}}>+ 사업자 추가</button>
            )}
          </div>
        </Card>
      ))}
    </div>
  );
}

function EditMembersPanel({ users, businesses, onRefresh, showToast }) {
  const [selId, setSelId] = useState(null);
  const [form, setForm] = useState({name:"",email:"",phone:"",password:"",memo:"",status:"approved"});
  const [saving, setSaving] = useState(false);
  // 사업자 편집
  const [newBizNo, setNewBizNo] = useState("");
  const [newBizName, setNewBizName] = useState("");
  const [addingBiz, setAddingBiz] = useState(false);

  const selUser = users.find(u=>u.id===selId);
  const selectUser = (u) => {
    setSelId(u.id);
    setForm({ name:u.name||"", email:u.email||"", phone:u.phone||"", password:"", memo:u.memo||"", status:u.status||"approved" });
    setNewBizNo(""); setNewBizName(""); setAddingBiz(false);
  };

  const save = () => {
    if(!selId) return;
    setSaving(true);
    db.updateUser(selId, form)
      .then(()=>{ onRefresh(); showToast("저장되었습니다."); setSaving(false); })
      .catch(()=>{ showToast("저장 중 오류가 발생했습니다."); setSaving(false); });
  };

  const addBiz = async () => {
    if(!newBizNo.trim()){ showToast("사업자번호를 입력하세요."); return; }
    const trimmed = newBizNo.trim();
    const bizName = newBizName.trim() || trimmed;
    try {
      // businesses 테이블에 등록
      await db.addBiz(trimmed, { name:bizName, type:"개인", representative:form.name });
      // user_businesses 연결
      const current = selUser?.businesses || [];
      if(!current.includes(trimmed)){
        await db.setBizList(selId, [...current, trimmed]);
      }
      setNewBizNo(""); setNewBizName(""); setAddingBiz(false);
      onRefresh(); showToast("사업자번호가 추가되었습니다.");
    } catch(e) { showToast("오류가 발생했습니다."); }
  };

  const removeBiz = async (biz) => {
    const current = selUser?.businesses || [];
    await db.setBizList(selId, current.filter(b=>b!==biz)).catch(()=>{});
    onRefresh();
  };

  const statusColors = {approved:[T.green,T.greenLight],pending:[T.orange,T.orangeLight],rejected:[T.red,T.redLight]};
  const selStyle = {padding:"10px 14px",borderRadius:T.radiusSm,border:`1.5px solid ${T.border}`,background:"rgba(255,255,255,0.9)",color:T.text,fontSize:"14px",fontFamily:T.font,outline:"none",width:"100%",boxSizing:"border-box"};

  return (
    <div>
      <h2 style={{color:T.text,fontSize:"20px",fontWeight:"700",margin:"0 0 6px",letterSpacing:"-0.7px"}}>회원 정보 수정</h2>
      <p style={{color:T.textSub,fontSize:"14px",margin:"0 0 20px"}}>회원을 선택한 후 정보를 수정하세요.</p>
      <div style={{display:"grid",gridTemplateColumns:"260px 1fr",gap:"20px",alignItems:"start"}}>
        {/* 회원 목록 */}
        <Card style={{padding:"16px",maxHeight:"600px",overflow:"auto"}}>
          <p style={{color:T.textSub,fontSize:"11px",fontWeight:"700",textTransform:"uppercase",letterSpacing:"0.8px",margin:"0 0 12px"}}>회원 목록</p>
          {users.length===0&&<p style={{color:T.textMuted,fontSize:"13px"}}>가입자가 없습니다.</p>}
          {users.map(u=>{
            const [sc,sb]=statusColors[u.status]||[T.textMuted,"transparent"];
            return (
              <div key={u.id} onClick={()=>selectUser(u)} style={{
                padding:"11px 14px",borderRadius:T.radiusSm,cursor:"pointer",marginBottom:"4px",
                background:selId===u.id?T.blueLight:"rgba(0,0,0,0.02)",
                border:`1.5px solid ${selId===u.id?T.blue:T.border}`,transition:"all 0.15s",
              }}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <span style={{color:T.text,fontWeight:"600",fontSize:"14px"}}>{u.name}</span>
                  <span style={{background:sb,color:sc,fontSize:"10px",fontWeight:"700",padding:"2px 8px",borderRadius:"20px"}}>{u.status==="approved"?"승인":u.status==="pending"?"대기":"거절"}</span>
                </div>
                <p style={{color:T.textSub,fontSize:"12px",margin:"2px 0 0"}}>{u.email}</p>
              </div>
            );
          })}
        </Card>

        {/* 편집 폼 */}
        {selUser ? (
          <Card style={{padding:"28px"}}>
            <p style={{color:T.textSub,fontSize:"11px",fontWeight:"700",textTransform:"uppercase",letterSpacing:"0.8px",margin:"0 0 18px"}}>{selUser.name} 정보 수정</p>

            {/* 기본 정보 */}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px",marginBottom:"12px"}}>
              <div><Label>이름</Label><Input value={form.name} onChange={v=>setForm(x=>({...x,name:v}))}/></div>
              <div><Label>연락처</Label><Input value={form.phone} onChange={v=>setForm(x=>({...x,phone:v}))} placeholder="010-0000-0000"/></div>
            </div>
            <div style={{marginBottom:"12px"}}><Label>이메일</Label><Input value={form.email} onChange={v=>setForm(x=>({...x,email:v}))} type="email"/></div>
            <div style={{marginBottom:"12px"}}><Label>새 비밀번호 (변경 시만 입력)</Label><Input value={form.password} onChange={v=>setForm(x=>({...x,password:v}))} type="password" placeholder="변경하지 않으면 비워두세요"/></div>
            <div style={{marginBottom:"12px"}}>
              <Label>상태</Label>
              <select value={form.status} onChange={e=>setForm(x=>({...x,status:e.target.value}))} style={selStyle}>
                <option value="approved">승인</option>
                <option value="pending">대기</option>
                <option value="rejected">거절</option>
              </select>
            </div>
            <div style={{marginBottom:"20px"}}><Label>메모</Label>
              <textarea value={form.memo} onChange={e=>setForm(x=>({...x,memo:e.target.value}))} placeholder="관리자 메모"
                style={{width:"100%",padding:"10px 14px",borderRadius:T.radiusSm,border:`1.5px solid ${T.border}`,background:"rgba(255,255,255,0.9)",color:T.text,fontSize:"14px",fontFamily:T.font,outline:"none",resize:"none",height:"72px",boxSizing:"border-box"}}/>
            </div>
            <div style={{display:"flex",gap:"10px",marginBottom:"28px"}}>
              <Btn onClick={save} disabled={saving}>{saving?"저장 중…":"기본정보 저장"}</Btn>
              <Btn onClick={()=>setSelId(null)} variant="secondary">취소</Btn>
            </div>

            {/* 사업자번호 관리 */}
            <div style={{borderTop:`1px solid ${T.border}`,paddingTop:"20px"}}>
              <p style={{color:T.textSub,fontSize:"11px",fontWeight:"700",textTransform:"uppercase",letterSpacing:"0.8px",margin:"0 0 12px"}}>연결된 사업자번호</p>
              <div style={{display:"flex",flexWrap:"wrap",gap:"6px",marginBottom:"10px"}}>
                {(selUser.businesses||[]).length===0&&<span style={{color:T.textMuted,fontSize:"13px"}}>등록된 사업자번호 없음</span>}
                {(selUser.businesses||[]).map(biz=>(
                  <span key={biz} style={{display:"inline-flex",alignItems:"center",gap:"5px",background:T.blueLight,color:T.blue,fontSize:"12px",fontWeight:"600",padding:"5px 12px",borderRadius:"20px"}}>
                    <span>{businesses[biz]?.name||biz}</span>
                    <span style={{color:T.textMuted,fontSize:"11px"}}>({biz})</span>
                    <button onClick={()=>removeBiz(biz)} style={{background:"none",border:"none",color:T.red,cursor:"pointer",padding:"0 0 0 4px",display:"inline-flex",alignItems:"center",fontSize:"14px",lineHeight:1}}>×</button>
                  </span>
                ))}
              </div>
              {addingBiz ? (
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr auto auto",gap:"8px",alignItems:"end"}}>
                  <div><Label>사업자번호 *</Label><Input value={newBizNo} onChange={setNewBizNo} placeholder="000-00-00000"/></div>
                  <div><Label>상호명 (선택)</Label><Input value={newBizName} onChange={setNewBizName} placeholder="(주)예시"/></div>
                  <Btn onClick={addBiz} size="sm">추가</Btn>
                  <Btn onClick={()=>{setAddingBiz(false);setNewBizNo("");setNewBizName("");}} variant="secondary" size="sm">취소</Btn>
                </div>
              ) : (
                <button onClick={()=>setAddingBiz(true)} style={{background:"none",border:`1px dashed ${T.borderStrong}`,color:T.blue,padding:"6px 16px",borderRadius:"8px",cursor:"pointer",fontSize:"13px",fontFamily:T.font,fontWeight:"600"}}>
                  + 사업자번호 추가
                </button>
              )}
            </div>
          </Card>
        ) : (
          <Card style={{padding:"60px",textAlign:"center"}}>
            <p style={{color:T.textMuted,fontSize:"14px"}}>좌측에서 수정할 회원을 선택하세요.</p>
          </Card>
        )}
      </div>
    </div>
  );
}

function BizPanel({ businesses, onRefresh, showToast }) {
  const [f,setF]=useState({bizNo:"",name:"",type:"법인",representative:""});
  const add=()=>{ if(!f.bizNo||!f.name){showToast("사업자번호와 상호는 필수입니다.");return;} db.addBiz(f.bizNo,{name:f.name,type:f.type,representative:f.representative}).then(()=>{ setF({bizNo:"",name:"",type:"법인",representative:""}); onRefresh(); showToast("등록되었습니다."); }).catch(()=>showToast("오류가 발생했습니다.")); };
  const selStyle={padding:"10px 14px",borderRadius:T.radiusSm,border:`1.5px solid ${T.border}`,background:"rgba(255,255,255,0.9)",color:T.text,fontSize:"14px",fontFamily:T.font,outline:"none",width:"100%",boxSizing:"border-box"};
  return (
    <div>
      <h2 style={{color:T.text,fontSize:"20px",fontWeight:"700",margin:"0 0 20px",letterSpacing:"-0.7px"}}>사업자 관리</h2>
      <Card style={{padding:"24px",marginBottom:"24px"}}>
        <p style={{color:T.textSub,fontSize:"11px",fontWeight:"700",textTransform:"uppercase",letterSpacing:"0.8px",margin:"0 0 14px"}}>신규 등록</p>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 100px 1fr",gap:"10px",alignItems:"end",marginBottom:"14px"}}>
          <div><Label>사업자번호 *</Label><Input value={f.bizNo} onChange={v=>setF(x=>({...x,bizNo:v}))} placeholder="000-00-00000"/></div>
          <div><Label>상호 *</Label><Input value={f.name} onChange={v=>setF(x=>({...x,name:v}))} placeholder="(주)한국상사"/></div>
          <div><Label>구분</Label><select style={selStyle} value={f.type} onChange={e=>setF(x=>({...x,type:e.target.value}))}><option>법인</option><option>개인</option></select></div>
          <div><Label>대표자</Label><Input value={f.representative} onChange={v=>setF(x=>({...x,representative:v}))} placeholder="홍길동"/></div>
        </div>
        <Btn onClick={add}>등록</Btn>
      </Card>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:"10px"}}>
        {Object.entries(businesses).map(([no,info])=>(
          <Card key={no} style={{padding:"18px"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div>
                <p style={{color:T.text,fontWeight:"700",fontSize:"14px",margin:"0 0 4px",letterSpacing:"-0.4px"}}>{info.name}</p>
                <p style={{color:T.textSub,fontSize:"12px",margin:"0 0 2px"}}>{no}</p>
                <p style={{color:T.textMuted,fontSize:"11px",margin:0}}>대표자: {info.representative||"-"}</p>
              </div>
              <Pill label={info.type} color={info.type==="법인"?T.blue:T.green} bg={info.type==="법인"?T.blueLight:T.greenLight}/>
            </div>
          </Card>
        ))}
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────
   카카오톡 채널 플로팅 버튼
───────────────────────────────────────── */
function KakaoFloat() {
  const isMobile = useIsMobile();
  const [hovered, setHovered] = useState(false);
  const [tooltip, setTooltip] = useState(true);
  useEffect(()=>{ const t=setTimeout(()=>setTooltip(false),3000); return()=>clearTimeout(t); },[]);

  // 모바일: 탭바(60px) + 여유(12px), PC: 28px
  const bottomPos = isMobile ? "calc(60px + env(safe-area-inset-bottom) + 12px)" : "28px";
  const rightPos  = isMobile ? "16px" : "28px";

  return (
    <div style={{position:"fixed",bottom:bottomPos,right:rightPos,zIndex:9999,display:"flex",flexDirection:"column",alignItems:"flex-end",gap:"8px"}} className="no-print">

      {/* 말풍선 툴팁 */}
      {(hovered||tooltip)&&(
        <div style={{
          background:"#FEE500",color:"#3A1D1D",
          fontSize:"13px",fontWeight:"700",
          padding:"9px 14px",borderRadius:"12px",
          boxShadow:"0 4px 16px rgba(0,0,0,0.15)",
          whiteSpace:"nowrap",letterSpacing:"-0.2px",
          position:"relative",
          animation:"fadeIn 0.2s ease",
        }}>
          담당자에게 바로 문의하기
          {/* 말풍선 꼬리 */}
          <div style={{
            position:"absolute",bottom:"-7px",right:"22px",
            width:"14px",height:"8px",
            background:"#FEE500",
            clipPath:"polygon(0 0, 100% 0, 50% 100%)",
          }}/>
        </div>
      )}

      {/* 카카오 버튼 */}
      <a
        href="http://pf.kakao.com/_QwKQn/chat"
        target="_blank"
        rel="noopener noreferrer"
        onMouseEnter={()=>setHovered(true)}
        onMouseLeave={()=>setHovered(false)}
        style={{
          display:"flex",alignItems:"center",justifyContent:"center",
          width:"56px",height:"56px",borderRadius:"50%",
          background:"#FEE500",
          boxShadow: hovered
            ? "0 8px 28px rgba(254,229,0,0.55), 0 2px 8px rgba(0,0,0,0.15)"
            : "0 4px 16px rgba(254,229,0,0.4), 0 2px 6px rgba(0,0,0,0.12)",
          transform: hovered ? "scale(1.08)" : "scale(1)",
          transition:"all 0.2s cubic-bezier(0.34,1.56,0.64,1)",
          textDecoration:"none",
        }}
      >
        {/* 카카오톡 공식 아이콘 SVG */}
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none">
          <path d="M12 3C7.03 3 3 6.36 3 10.5c0 2.67 1.7 5.02 4.27 6.37l-.97 3.6 4.18-2.77c.48.07.97.1 1.52.1 4.97 0 9-3.36 9-7.5S16.97 3 12 3z" fill="#3A1D1D"/>
        </svg>
      </a>

      <style>{`
        @keyframes fadeIn { from{opacity:0;transform:translateY(4px)} to{opacity:1;transform:translateY(0)} }
      `}</style>
    </div>
  );
}

/* ─────────────────────────────────────────
   메인 앱
───────────────────────────────────────── */
export default function App() {
  const [page,setPage]=useState("login");
  const [user,setUser]=useState(null);
  const login=(u)=>{ setUser(u); setPage(u.role==="admin"?"admin":"customer"); };
  const logout=()=>{ setUser(null); setPage("login"); };
  return (
    <div style={{fontFamily:T.font}}>
      {page==="login"&&<Login onLogin={login} onGo={setPage}/>}
      {page==="register"&&<Register onGo={setPage}/>}
      {page==="customer"&&<CustomerDash user={user} onLogout={logout}/>}
      {page==="admin"&&<AdminDash user={user} onLogout={logout}/>}
      <KakaoFloat/>
    </div>
  );
}
