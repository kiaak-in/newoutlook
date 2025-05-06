// function-file.js - 이벤트 핸들러
Office.onReady(() => {
  // Office.js가 준비되었을 때 실행
});

// 전역 함수로 등록하여 매니페스트에서 이벤트 핸들러로 사용
window.onNewMessageComposeHandler = (event) => {
  addDefaultSignature(event);
};

window.onReplyMessageComposeHandler = (event) => {
  addReplySignature(event);
};

window.onForwardMessageComposeHandler = (event) => {
  addForwardSignature(event);
};

// 새 메시지에 기본 서명 추가
function addDefaultSignature(event) {
  Office.context.mailbox.item.body.getAsync(
    Office.CoercionType.Html,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        let bodyContent = result.value;

        // 서명이 이미 있는지 확인
        if (!bodyContent.includes("<!-- 서명 ID: ")) {
          // 사용자 정보 가져오기
          const userProfile = Office.context.mailbox.userProfile;
          const userName = userProfile.displayName;
          const userEmail = userProfile.emailAddress;

          // 회사 정보 (실제로는 서버나 설정에서 가져와야 함)
          const companyInfo = {
            name: "회사명",
            address: "서울시 강남구 테헤란로 123",
            phone: "02-123-4567",
            website: "https://www.company.com",
          };

          // HTML 서명 생성
          const signature = generateSignatureHtml(
            userName,
            userEmail,
            companyInfo,
            "new"
          );

          // 현재 본문에 서명 추가
          Office.context.mailbox.item.body.setAsync(
            bodyContent + signature,
            { coercionType: Office.CoercionType.Html },
            (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("서명이 성공적으로 추가되었습니다.");
              } else {
                console.error("서명 추가 실패:", result.error);
              }
              event.completed();
            }
          );
        } else {
          // 이미 서명이 있음
          console.log("이미 서명이 있습니다.");
          event.completed();
        }
      } else {
        console.error("본문 가져오기 실패:", result.error);
        event.completed();
      }
    }
  );
}

// 답장 메시지에 서명 추가
function addReplySignature(event) {
  // 답장용 서명 추가 로직 (간소화된 서명 등)
  Office.context.mailbox.item.body.getAsync(
    Office.CoercionType.Html,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        let bodyContent = result.value;

        // 서명이 이미 있는지 확인
        if (!bodyContent.includes("<!-- 서명 ID: ")) {
          // 사용자 정보 가져오기
          const userProfile = Office.context.mailbox.userProfile;
          const userName = userProfile.displayName;
          const userEmail = userProfile.emailAddress;

          // 회사 정보
          const companyInfo = {
            name: "회사명",
            address: "서울시 강남구 테헤란로 123",
            phone: "02-123-4567",
            website: "https://www.company.com",
          };

          // 답장용 간소화된 서명 생성
          const signature = generateSignatureHtml(
            userName,
            userEmail,
            companyInfo,
            "reply"
          );

          // 본문에 서명 추가 (일반적으로 커서 위치에 추가)
          Office.context.mailbox.item.body.setAsync(
            bodyContent + signature,
            { coercionType: Office.CoercionType.Html },
            (result) => {
              event.completed();
            }
          );
        } else {
          event.completed();
        }
      } else {
        event.completed();
      }
    }
  );
}

// 전달 메시지에 서명 추가
function addForwardSignature(event) {
  // 전달용 서명 추가 로직 (기본 서명과 유사)
  addReplySignature(event); // 일반적으로 답장과 동일한 서명 사용
}

// HTML 서명 생성 함수
function generateSignatureHtml(name, email, company, type) {
  // 고유 ID 생성 (서명 중복 추가 방지용)
  const signatureId = Date.now().toString();

  // 서명 유형에 따라 다른 템플릿 사용
  let signatureHtml = "";

  if (type === "new") {
    // 새 메시지용 전체 서명
    signatureHtml = `
        <br/><br/>
        <!-- 서명 ID: ${signatureId} -->
        <div style="font-family: 'Malgun Gothic', Arial, sans-serif; font-size: 10pt; color: #333333; border-top: 1px solid #cccccc; padding-top: 10px; margin-top: 10px;">
          <div style="margin-bottom: 10px;">
            <strong style="font-size: 12pt; color: #004b8d;">${name}</strong>
            <br/>
            <span style="color: #666666;">회사명</span>
          </div>
          <div style="margin-bottom: 5px;">
            <span>Email: <a href="mailto:${email}" style="color: #004b8d; text-decoration: none;">${email}</a></span>
            <br/>
            <span>Tel: ${company.phone}</span>
          </div>
          <div style="margin-top: 10px;">
            <a href="${company.website}" style="color: #004b8d; text-decoration: none;">${company.website}</a>
            <br/>
            <span style="font-size: 9pt; color: #999999;">${company.address}</span>
          </div>
          <div style="margin-top: 10px;">
            <img src="https://your-domain.com/assets/logo.png" alt="${company.name}" style="max-width: 150px; height: auto;" />
          </div>
          <div style="font-size: 8pt; color: #999999; margin-top: 10px;">
            본 메일은 발신전용입니다. 궁금하신 사항은 고객센터로 문의해 주시기 바랍니다.
          </div>
        </div>
      `;
  } else {
    // 답장/전달용 간소화된 서명
    signatureHtml = `
        <br/>
        <!-- 서명 ID: ${signatureId} -->
        <div style="font-family: 'Malgun Gothic', Arial, sans-serif; font-size: 10pt; color: #333333;">
          <div>
            <strong>${name}</strong> | ${company.name} | <a href="mailto:${email}" style="color: #004b8d; text-decoration: none;">${email}</a> | Tel: ${company.phone}
          </div>
        </div>
      `;
  }

  return signatureHtml;
}
