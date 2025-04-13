(function () {
  "use strict";

  // Office 초기화
  Office.onReady(function (info) {
    // Office.js가 로드되면 실행
    if (info.host === Office.HostType.Outlook) {
      // 명령 핸들러 등록
      Office.actions.associate("showPopup", showPopup);
    }
  });

  // 팝업을 표시하는 함수
  function showPopup(event) {
    // 팝업 대화상자 옵션
    var dialogOptions = {
      width: 30,
      height: 30,
      displayInIframe: true,
    };

    // 팝업 대화상자 열기
    Office.context.ui.displayDialogAsync(
      "https://kiaak-in.github.io/newoutlook/popup.html",
      dialogOptions,
      function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          // 오류 처리
          console.error(
            "대화상자를 열지 못했습니다. 오류: " + result.error.message
          );
          event.completed();
        } else {
          // 대화상자 객체 저장
          var dialog = result.value;

          // 대화상자 이벤트 핸들러
          dialog.addEventHandler(
            Office.EventType.DialogMessageReceived,
            function (arg) {
              // 대화상자에서 메시지 수신 시 처리
              console.log("대화상자에서 메시지 수신: " + arg.message);
              dialog.close();
              event.completed();
            }
          );

          // 대화상자가 닫힐 때 이벤트 핸들러
          dialog.addEventHandler(
            Office.EventType.DialogEventReceived,
            function (arg) {
              // 대화상자가 닫히면 이벤트 완료 처리
              console.log("대화상자 닫힘 이벤트: " + arg.message);
              event.completed();
            }
          );
        }
      }
    );
  }
})();
