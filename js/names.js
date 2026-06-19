// Name detection and short-name resolution.
// Vietnamese & Korean: family name first → short name = last 2 given names (after family name).
// Japanese: family name first → short name = last word (single given name).
// All other names: short name = first word (Western given-name-first order).

const VIET_SURNAMES = new Set([
  'nguyen','le','tran','pham','hoang','huynh','phan','vu','vo','dang',
  'bui','do','ho','ngo','duong','ly','dao','dinh','mai','trieu','luong',
  'trinh','nghiem','luu','truong','lam','to','thai','ta','ky','khuc'
]);
function isViet(n){ return VIET_SURNAMES.has(n.trim().split(/\s+/)[0].toLowerCase()); }

const JAPANESE_SURNAMES = new Set([
  'araki','sato','suzuki','takahashi','tanaka','watanabe','ito','yamamoto',
  'nakamura','kobayashi','kato','yoshida','yamada','sasaki','yamaguchi',
  'matsumoto','inoue','kimura','hayashi','shimizu','yamazaki','mori','abe',
  'ikeda','hashimoto','yamashita','ishikawa','nakajima','maeda','fujita',
  'ogawa','okamoto','matsuda','nishimura','fujii','goto','okada','hasegawa',
  'murakami','kondo','ishii','saito','aoki','fujiwara','nakano','ueda',
  'ogata','ishida','nishida','miyamoto','murata','hara','ono'
]);
function isJapanese(n){ return JAPANESE_SURNAMES.has(n.trim().split(/\s+/)[0].toLowerCase()); }

const KOREAN_SURNAMES = new Set([
  'kim','lee','park','choi','jung','kang','cho','yoon','jang','lim',
  'han','oh','seo','shin','kwon','hwang','ahn','song','hong','jeon',
  'yu','ryu','noh','heo','sim','baek','moon','yang','nam',
  'ha','joo','ku','yoo','min','chun','bae','cha','roh','pi'
]);
function isKorean(n){ return KOREAN_SURNAMES.has(n.trim().split(/\s+/)[0].toLowerCase()); }

function parseBracketed(raw){
  const m = raw.trim().match(/^(.+?)\s*\((.+?)\)\s*$/);
  return m ? { fullName: m[1].trim(), nickname: m[2].trim() } : { fullName: raw.trim(), nickname: '' };
}

function displayName(s){ return s.nickname ? `${s.fullName} (${s.nickname})` : s.fullName; }

function shortName(s){
  if(s.nickname) return s.nickname;
  const parts = s.fullName.trim().split(/\s+/);
  if(s.viet || s.korean){
    const noFam = parts.slice(1);
    return noFam.length >= 2 ? noFam.slice(-2).join(' ') : noFam.join(' ') || parts[0];
  }
  if(s.jap){
    return parts[parts.length - 1];
  }
  return parts[0];
}
