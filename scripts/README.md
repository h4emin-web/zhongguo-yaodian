# 로컬 자동정산 도구

ERP 환율 조회는 ECOUNT 공식 OpenAPI에 `수입비용전표검색` 조회 API가 없어, 로컬 Chrome 자동화로 처리합니다.

## 최초 1회

```powershell
npm install
```

## 환율 조회

```powershell
$env:ECOUNT_COM_CODE="23263"
$env:ECOUNT_USER_ID="강해민"
$env:ECOUNT_PASSWORD="비밀번호"
$env:PO_NO="Z26-00577"
npm run ecount:rate
```

성공하면 JSON으로 환율이 출력됩니다.

## 사이트에서 ERP 조회 버튼까지 사용

```powershell
$env:ECOUNT_COM_CODE="23263"
$env:ECOUNT_USER_ID="강해민"
$env:ECOUNT_PASSWORD="비밀번호"
npm run local
```

브라우저에서 `http://localhost:4173`을 열면 자동정산의 `ERP 조회` 버튼이 로컬 ERP 자동화와 연결됩니다.
