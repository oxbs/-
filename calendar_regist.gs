
//�R���e�i�o�C���h�X�N���v�g�i�X�v���b�h�V�[�g�ƘA�g�j

//�C�x���g�n���h���F�C�x���g�������ɏ�������ionOpen�̓X�v���b�h�V�[�g���J���Ǝ��s�j
function onOpen() { 

  let ui = SpreadsheetApp.getUi()

  //�u�ǉ����j���[�v�Ƃ������j���[���X�v���b�h�V�[�g�ɒǉ������
  let menu = ui.createMenu("�ǉ����j���["); 

  //�u�J�����_�[�o�^�v�Ƃ����A�C�e������ݒ�AcalenderAdd�̊֐������s
  menu.addItem("�J�����_�[�o�^", "calenderAdd"); 

  //�u�N���A�v�Ƃ����A�C�e������ݒ�AclearCell�̊֐������s
  menu.addItem("�N���A", "clearCell");

  //��ʏ�̃��j���[�Ƃ��Ēǉ����邽�߂ɕK�v�ȏ���
  menu.addToUi(); 
}

function calenderAdd() {

  var data = SpreadsheetApp.getActiveSheet().getDataRange().getValues();

  //1�s�ڂ͍��ږ��A2�s�ڂ̓T���v�����͍s�Ȃ̂ō폜
  data.shift();
  data.shift();

  for (var row of data){
    var title = row[0];
    var date = row[1];
    var startTime = row[2];
    var endTime = row[3];

    //new Date�ŐV�������t�I�u�W�F�N�g�Ƃ��Ē�`���Ȃ��ƁA���̓��t�ϐ��Ƌ���������ۂ�
    var startDate = new Date(date); 

    //���t�Ɏ��ԏ���t�^
    startDate.setHours(startTime.getHours()); 

    //���t�ɕ�����t�^�@�˃J�����_�[�o�^�ł�������`���ɂȂ�
    startDate.setMinutes(startTime.getMinutes()); 

    var endDate = new Date(date);
    endDate.setHours(endTime.getHours());
    endDate.setMinutes(endTime.getMinutes());

    //�I�v�V�����Ƃ��āu�����v�u�ꏊ�v���ڂ��擾
    var option = {
      description: row[4],
      location: row[5]
    }

    //Google�J�����_�[���Ăяo���i�����ɂ͎��g��ID�iGoogle���[���A�h���X�j��ݒ�j
    let calender = CalendarApp.getCalendarById("oxbs2005@gmail.com"); 

    //�J�����_�[�ɃC�x���g�o�^
    calender.createEvent( 
      title,
      startDate,
      endDate,
      option
    );

  }

  //�u���E�U��Ƀ|�b�v�A�b�v���b�Z�[�W���o��
  Browser.msgBox("�J�����_�[�ɓo�^���܂���");
}

function clearCell(){

  let rSheet = SpreadsheetApp.getActiveSheet();
  let lastRow = rSheet.getLastRow(); 

  //�폜�͈͂̐ݒ�
  var clearData = rSheet.getRange(3,1,lastRow-2,6);

  //�Z���̓��͓��e�������폜�i�����Ȃǂ͕ς��Ȃ��j
  clearData.clearContent();
}

