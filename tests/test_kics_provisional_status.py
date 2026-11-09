from parse_annual import _extract_kics


def test_explicit_provisional_kics_does_not_fill_final_columns(tmp_path):
    path=tmp_path/'annual.xml'
    path.write_text('''<DOCUMENT>지급여력비율
      <TABLE><TR><TD>구분</TD><TD>2025년말</TD></TR>
      <TR><TD>지급여력금액(A)</TD><TD>21,555,174</TD></TR>
      <TR><TD>지급여력기준금액(B)</TD><TD>9,895,449</TD></TR>
      <TR><TD>지급여력비율(A/B)</TD><TD>217.8</TD></TR></TABLE>
      <P>당기말은 외부검증이 완료되지 않은 상태에서 작성된 잠정치입니다.</P>
      </DOCUMENT>''',encoding='utf-8')
    result=_extract_kics(path)
    assert result[92]==2.178 and result[94]==9895449
    assert not {95,96,97}.intersection(result)
