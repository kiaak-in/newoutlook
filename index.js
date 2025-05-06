// index.js - 애드인 주요 기능
let signatures = []; // 사용 가능한 서명 배열

// 애드인 초기화
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // 서명 목록 가져오기 (실제로는 서버에서 가져옴)
    loadSignatures();

    // UI 이벤트 핸들러 설정
    document.getElementById("btnApplySignature").onclick =
      applySelectedSignature;
  }
});

// 서명 목록 로드 (서버에서 가져오는 것을 시뮬레이션)
function loadSignatures() {
  // 실제 구현에서는 API를 통해 서버에서 사용자별 서명 가져오기
  signatures = [
    {
      id: "signature1",
      name: "기본 서명",
      html: generateDefaultSignature(),
    },
    {
      id: "signature2",
      name: "간소화 서명",
      html: generateSimpleSignature(),
    },
    {
      id: "signature3",
      name: "영문 서명",
      html: generateEnglishSignature(),
    },
  ];

  // 서명 목록 UI 업데이트
  updateSignatureList();
}

// 서명 목록 UI 업데이트
function updateSignatureList() {
  const signatureList = document.getElementById("signatureList");
  signatureList.innerHTML = "";

  signatures.forEach((signature) => {
    const option = document.createElement("option");
    option.value = signature.id;
    option.text = signature.name;
    signatureList.appendChild(option);
  });

  // 미리보기 업데이트
  updateSignaturePreview();
}

// 서명 미리보기 업데이트
function updateSignaturePreview() {
  const selectedId = document.getElementById("signatureList").value;
  const previewDiv = document.getElementById("signaturePreview");

  const selectedSignature = signatures.find((sig) => sig.id === selectedId);
  if (selectedSignature) {
    previewDiv.innerHTML = selectedSignature.html;
  } else {
    previewDiv.innerHTML = "<p>선택된 서명이 없습니다.</p>";
  }
}

// 선택한 서명 적용
function applySelectedSignature() {
  const selectedId = document.getElementById("signatureList").value;
  const selectedSignature = signatures.find((sig) => sig.id === selectedId);

  if (selectedSignature) {
    // 현재 메일 본문에 서명 추가
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          let body = result.value;

          // 기존 서명 제거 (서명 ID 태그로 식별)
          const signatureRegex =
            /<!-- 서명 ID: .*?-->([\s\S]*?)(?=<\/body>|$)/g;
          body = body.replace(signatureRegex, "");

          // 새 서명 추가
          Office.context.mailbox.item.body.setAsync(
            body + selectedSignature.html,
            { coercionType: Office.CoercionType.Html },
            (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                // 성공 메시지 표시
                document.getElementById("statusMessage").textContent =
                  "서명이 적용되었습니다.";
                setTimeout(() => {
                  document.getElementById("statusMessage").textContent = "";
                }, 3000);
              } else {
                document.getElementById("statusMessage").textContent =
                  "서명 적용 실패: " + result.error.message;
              }
            }
          );
        }
      }
    );
  }
}

// 기본 서명 HTML 생성 (예시)
function generateDefaultSignature() {
  const userProfile = Office.context.mailbox.userProfile;
  return `
    <!-- 서명 ID: ${Date.now()} -->
    <div style="font-family: 'Malgun Gothic', Arial, sans-serif; font-size: 10pt; color: #333333; border-top: 1px solid #cccccc; padding-top: 10px; margin-top: 10px;">
      <div style="margin-bottom: 10px;">
        <strong style="font-size: 12pt; color: #004b8d;">${
          userProfile.displayName
        }</strong>
        <br/>
        <span style="color: #666666;">팀장 | 개발팀</span>
      </div>
      <div style="margin-bottom: 5px;">
        <span>Email: <a href="mailto:${
          userProfile.emailAddress
        }" style="color: #004b8d; text-decoration: none;">${
    userProfile.emailAddress
  }</a></span>
        <br/>
        <span>Tel: 02-123-4567</span>
      </div>
      <div style="margin-top: 10px;">
        <a href="https://www.company.com" style="color: #004b8d; text-decoration: none;">www.company.com</a>
        <br/>
        <span style="font-size: 9pt; color: #999999;">서울시 강남구 테헤란로 123</span>
      </div>
      <div style="margin-top: 10px;">
        <img src="https://your-domain.com/assets/logo.png" alt="회사로고" style="max-width: 150px; height: auto;" />
      </div>
    </div>
  `;
}

// 간소화된 서명 HTML 생성 (예시)
function generateSimpleSignature() {
  const userProfile = Office.context.mailbox.userProfile;
  return `
    <!-- 서명 ID: ${Date.now()} -->
    <div style="font-family: 'Malgun Gothic', Arial, sans-serif; font-size: 10pt; color: #333333;">
      <div>
        <strong>${
          userProfile.displayName
        }</strong> | 회사명 | <a href="mailto:${
    userProfile.emailAddress
  }" style="color: #004b8d; text-decoration: none;">${
    userProfile.emailAddress
  }</a> | Tel: 02-123-4567
      </div>
    </div>
  `;
}

// 영문 서명 HTML 생성 (예시)
function generateEnglishSignature() {
  const userProfile = Office.context.mailbox.userProfile;
  return `
    <!-- 서명 ID: ${Date.now()} -->
    <div style="font-family: Arial, sans-serif; font-size: 10pt; color: #333333; border-top: 1px solid #cccccc; padding-top: 10px; margin-top: 10px;">
      <div style="margin-bottom: 10px;">
        <strong style="font-size: 12pt; color: #004b8d;">${
          userProfile.displayName
        }</strong>
        <br/>
        <span style="color: #666666;">Team Leader | Development Team</span>
      </div>
      <div style="margin-bottom: 5px;">
        <span>Email: <a href="mailto:${
          userProfile.emailAddress
        }" style="color: #004b8d; text-decoration: none;">${
    userProfile.emailAddress
  }</a></span>
        <br/>
        <span>Tel: +82-2-123-4567</span>
      </div>
      <div style="margin-top: 10px;">
        <a href="https://www.company.com" style="color: #004b8d; text-decoration: none;">www.company.com</a>
        <br/>
        <span style="font-size: 9pt; color: #999999;">123 Teheran-ro, Gangnam-gu, Seoul, Korea</span>
      </div>
    </div>
  `;
}
