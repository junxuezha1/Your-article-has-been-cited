from extract_references import (
    _guess_authors_from_paragraphs,
    _guess_title_from_paragraphs,
    _parse_single_reference,
)


def test_article_metadata_keeps_title_continuation_out_of_authors():
    paragraphs = [
        "论合作社资本的功能",
        "——以资本与惠顾的功能比较为视角",
        "张素华，吴亦伟",
        "摘要：合作社资本制度需要兼顾成员权益。",
    ]

    assert _guess_title_from_paragraphs(paragraphs) == "论合作社资本的功能——以资本与惠顾的功能比较为视角"
    assert _guess_authors_from_paragraphs(paragraphs) == "张素华，吴亦伟"


def test_article_metadata_skips_affiliation_lines_before_authors():
    paragraphs = [
        "新时代党建引领乡村治理的实践逻辑及优化进路",
        "(南京师范大学马克思主义学院，南京，210023)",
        "吴建雄",
        "摘要：党建引领乡村治理需要优化组织机制。",
    ]

    assert _guess_title_from_paragraphs(paragraphs) == "新时代党建引领乡村治理的实践逻辑及优化进路"
    assert _guess_authors_from_paragraphs(paragraphs) == "吴建雄"


def test_article_metadata_accepts_two_character_author_name():
    paragraphs = [
        "DOI: 10.11817/j.issn. 1672-3104. 2026. 01.019",
        "央地财政分权何以影响区域协调发展？",
        "——一个政治经济学分析",
        "肖芸",
        "(中山大学马克思主义学院，广东广州，510275)",
        "摘要：央地财政分权如何影响区域协调发展是政治经济学研究中相对被忽略的议题。",
    ]

    assert _guess_title_from_paragraphs(paragraphs) == "央地财政分权何以影响区域协调发展？——一个政治经济学分析"
    assert _guess_authors_from_paragraphs(paragraphs) == "肖芸"


def test_article_metadata_skips_colon_subtitle_before_authors():
    paragraphs = [
        "DOI: 10.11817/j.issn. 1672-3104. 2026. 01.013",
        "数字时代的意识形态叙事：",
        "价值意蕴、潜在危机与化解策略",
        "曹清燕，欧露雯",
        "(中南大学马克思主义学院，湖南长沙，410083)",
        "摘要：意识形态工作是党的一项极端重要的工作。",
    ]

    assert _guess_title_from_paragraphs(paragraphs) == "数字时代的意识形态叙事： 价值意蕴、潜在危机与化解策略"
    assert _guess_authors_from_paragraphs(paragraphs) == "曹清燕，欧露雯"


def test_article_metadata_keeps_long_title_subtitle_together():
    paragraphs = [
        "央地财政分权何以影响区域协调发展？",
        "——一个政治经济学分析",
        "张晏1，王永钦2",
        "摘要：财政分权影响区域协调发展。",
    ]

    assert _guess_title_from_paragraphs(paragraphs) == "央地财政分权何以影响区域协调发展？——一个政治经济学分析"
    assert _guess_authors_from_paragraphs(paragraphs) == "张晏，王永钦"


def test_reference_parser_accepts_chinese_authors_without_superscript_numbers():
    ref = _parse_single_reference("吕冰洋, 陈怡心, 詹静楠. 政府预算管理、征税行为与企业经营效率[J]. 经济研究, 2021, 56(8): 42-58.")

    assert ref["authors"] == "吕冰洋, 陈怡心, 詹静楠"
    assert ref["title"] == "政府预算管理、征税行为与企业经营效率"
    assert ref["journal"] == "经济研究"
