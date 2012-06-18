# coding:Shift_JIS

require 'spec_helper'

describe WorkBook, '��Shift_JIS�̃R�[�h����J���ꂽ�ꍇ' do
  subject{WorkBook.open('./spec/test_data/Book1.xlsx')}

  after do
    subject.close
  end

  describe '#open' do
    context '��Shift_JIS�ŏ����ꂽ�R�[�h����Ăяo����ăt�@�C�����J�����ɂɃG���R�[�h�𖾎��I�Ɏw�肵�Ȃ��Ă��A' do
      it '�V�[�g�����擾���镶���R�[�h��Shift_JIS�ɂȂ�܂��B' do
        subject.sheets[2].name.should == '���낢��ȃf�[�^'
      end

      it '�V�[�g��I�����镶����͐�������܂���B'  do
        subject['���낢��ȃf�[�^'].should equal subject.sheets[2]
      end

      it '�V�[�g���擾���郁�\�b�h�͐�������܂���B'  do
        subject.���낢��ȃf�[�^.should equal subject.sheets[2]
      end

      it '������f�[�^�̕����R�[�h��Shift_JIS�ɂȂ�܂��B' do
        subject.���낢��ȃf�[�^.A3.should == '����������'
      end
    end
  end
end