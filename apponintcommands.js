Office.onReady(function () {
  // Office가 초기화되었을 때 실행됩니다.
  Office.initialize = function (reason) {
    // 초기화 코드가 필요하다면 여기에 작성합니다.
  };
});

// 새 일정을 만드는 함수
function createNewAppointment(event) {
  // 새 일정 항목을 생성합니다.
  Office.context.mailbox.displayNewAppointmentForm({
    subject: "", // 제목 (빈 문자열로 설정)
    start: new Date(), // 시작 시간 (현재 시간)
    end: new Date(new Date().getTime() + 30 * 60 * 1000), // 종료 시간 (30분 후)
    location: "", // 위치
    // 필요시 추가 설정
    // attendees: [],          // 참석자 목록
    // resources: [],          // 자원 목록
    // body: {                 // 본문 내용
    //   type: Office.MailboxEnums.BodyType.Html,
    //   content: ''
    // }
  });

  // 이벤트 완료
  event.completed();
}

// Office.actions에 함수를 등록합니다.
Office.actions.associate("createNewAppointment", createNewAppointment);
