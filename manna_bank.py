# -*- coding: utf-8 -*-
"""
manna_bank.py — [오늘의 만나] 기본 자료집

AI 가 만들어 주는 것이 1순위이고, 이 파일은 AI 가 안 될 때(사용량 한도·서버 지연 등)
쓰는 예비 자료입니다. 그런데도 하루하루 다른 말씀과 예화가 나오도록,
구절 73편과 예화 71편을 서로 다른 주기로 맞물려 돌립니다.

  73 과 71 은 서로소(공약수가 1뿐)이므로
  (구절, 예화) 짝은 73 × 71 = 5,183일, 곧 약 14년 동안 한 번도 겹치지 않습니다.
  → 1년 365일이 모두 다른 조합입니다.

지켜야 할 원칙
  · 한글 본문은 모두 개역개정입니다.
  · 예화는 실제로 있었던 일(역사적 사건·실존 인물·성경 본문·작품 창작 배경)만 담습니다.
    지어낸 '어떤 성도의 이야기'는 넣지 않습니다.
"""

# ==============================================================================
# 구절 73편 — 개역개정 + NIV
# ==============================================================================
VERSES = [
    ("시편 23:1", "여호와는 나의 목자시니 내게 부족함이 없으리로다",
     "Ps 23:1", "The LORD is my shepherd, I lack nothing.", "목자 되신 주"),
    ("시편 1:1", "복 있는 사람은 악인들의 꾀를 따르지 아니하며 죄인들의 길에 서지 아니하며 "
     "오만한 자들의 자리에 앉지 아니하고",
     "Ps 1:1", "Blessed is the one who does not walk in step with the wicked or stand in the "
     "way that sinners take or sit in the company of mockers.", "복 있는 사람"),
    ("시편 46:1", "하나님은 우리의 피난처시요 힘이시니 환난 중에 만날 큰 도움이시라",
     "Ps 46:1", "God is our refuge and strength, an ever-present help in trouble.", "피난처"),
    ("시편 119:105", "주의 말씀은 내 발에 등이요 내 길에 빛이니이다",
     "Ps 119:105", "Your word is a lamp for my feet, a light on my path.", "말씀의 등불"),
    ("시편 121:1-2", "내가 산을 향하여 눈을 들리라 나의 도움이 어디서 올까 "
     "나의 도움은 천지를 지으신 여호와에게서로다",
     "Ps 121:1-2", "I lift up my eyes to the mountains—where does my help come from? "
     "My help comes from the LORD, the Maker of heaven and earth.", "도움이 어디서"),
    ("시편 37:5", "네 길을 여호와께 맡기라 그를 의지하면 그가 이루시고",
     "Ps 37:5", "Commit your way to the LORD; trust in him and he will do this.", "맡기는 삶"),
    ("시편 34:8", "너희는 여호와의 선하심을 맛보아 알지어다 그에게 피하는 자는 복이 있도다",
     "Ps 34:8", "Taste and see that the LORD is good; blessed is the one who takes refuge in him.",
     "맛보아 아는 은혜"),
    ("시편 51:10", "하나님이여 내 속에 정한 마음을 창조하시고 내 안에 정직한 영을 새롭게 하소서",
     "Ps 51:10", "Create in me a pure heart, O God, and renew a steadfast spirit within me.",
     "새롭게 하소서"),
    ("시편 55:22", "네 짐을 여호와께 맡기라 그가 너를 붙드시고 의인의 요동함을 영원히 허락하지 "
     "아니하시리로다",
     "Ps 55:22", "Cast your cares on the LORD and he will sustain you; he will never let the "
     "righteous be shaken.", "짐을 맡기라"),
    ("시편 103:2", "내 영혼아 여호와를 송축하며 그의 모든 은택을 잊지 말지어다",
     "Ps 103:2", "Praise the LORD, my soul, and forget not all his benefits.", "잊지 말 은택"),
    ("시편 126:5", "눈물을 흘리며 씨를 뿌리는 자는 기쁨으로 거두리로다",
     "Ps 126:5", "Those who sow with tears will reap with songs of joy.", "눈물의 씨앗"),
    ("시편 27:1", "여호와는 나의 빛이요 나의 구원이시니 내가 누구를 두려워하리요",
     "Ps 27:1", "The LORD is my light and my salvation—whom shall I fear?", "빛이요 구원"),
    ("시편 62:1", "나의 영혼이 잠잠히 하나님만 바람이여 나의 구원이 그에게서 나오는도다",
     "Ps 62:1", "Truly my soul finds rest in God; my salvation comes from him.", "잠잠히 바람"),
    ("시편 90:12", "우리에게 우리 날 계수함을 가르치사 지혜로운 마음을 얻게 하소서",
     "Ps 90:12", "Teach us to number our days, that we may gain a heart of wisdom.", "날을 세는 지혜"),
    ("시편 133:1", "보라 형제가 연합하여 동거함이 어찌 그리 선하고 아름다운고",
     "Ps 133:1", "How good and pleasant it is when God's people live together in unity!", "연합"),
    ("시편 139:23-24", "하나님이여 나를 살피사 내 마음을 아시며 나를 시험하사 내 뜻을 아옵소서",
     "Ps 139:23-24", "Search me, God, and know my heart; test me and know my anxious thoughts.",
     "나를 살피소서"),
    ("시편 42:1", "하나님이여 사슴이 시냇물을 찾기에 갈급함 같이 내 영혼이 주를 찾기에 갈급하니이다",
     "Ps 42:1", "As the deer pants for streams of water, so my soul pants for you, my God.",
     "갈급한 영혼"),
    ("시편 116:1", "여호와께서 내 음성과 내 간구를 들으시므로 내가 그를 사랑하는도다",
     "Ps 116:1", "I love the LORD, for he heard my voice; he heard my cry for mercy.", "들으시는 주"),
    ("잠언 3:5", "너는 마음을 다하여 여호와를 신뢰하고 네 명철을 의지하지 말라",
     "Pr 3:5", "Trust in the LORD with all your heart and lean not on your own understanding.",
     "신뢰"),
    ("잠언 3:6", "너는 범사에 그를 인정하라 그리하면 네 길을 지도하시리라",
     "Pr 3:6", "In all your ways submit to him, and he will make your paths straight.", "인정함"),
    ("잠언 4:23", "모든 지킬 만한 것 중에 더욱 네 마음을 지키라 생명의 근원이 이에서 남이니라",
     "Pr 4:23", "Above all else, guard your heart, for everything you do flows from it.", "마음 지킴"),
    ("잠언 16:9", "사람이 마음으로 자기의 길을 계획할지라도 그의 걸음을 인도하시는 이는 여호와시니라",
     "Pr 16:9", "In their hearts humans plan their course, but the LORD establishes their steps.",
     "걸음을 인도하심"),
    ("잠언 15:1", "유순한 대답은 분노를 쉬게 하여도 과격한 말은 노를 격동하느니라",
     "Pr 15:1", "A gentle answer turns away wrath, but a harsh word stirs up anger.", "유순한 대답"),
    ("잠언 17:22", "마음의 즐거움은 양약이라도 심령의 근심은 뼈를 마르게 하느니라",
     "Pr 17:22", "A cheerful heart is good medicine, but a crushed spirit dries up the bones.",
     "즐거운 마음"),
    ("잠언 22:6", "마땅히 행할 길을 아이에게 가르치라 그리하면 늙어도 그것을 떠나지 아니하리라",
     "Pr 22:6", "Start children off on the way they should go, and even when they are old they "
     "will not turn from it.", "다음 세대"),
    ("전도서 3:1", "범사에 기한이 있고 천하 만사가 다 때가 있나니",
     "Ecc 3:1", "There is a time for everything, and a season for every activity under the heavens.",
     "때가 있나니"),
    ("이사야 40:31", "오직 여호와를 앙망하는 자는 새 힘을 얻으리니 독수리가 날개치며 올라감 "
     "같을 것이요 달음박질하여도 곤비하지 아니하겠고 걸어가도 피곤하지 아니하리로다",
     "Isa 40:31", "But those who hope in the LORD will renew their strength. They will soar on "
     "wings like eagles; they will run and not grow weary, they will walk and not be faint.",
     "새 힘"),
    ("이사야 41:10", "두려워하지 말라 내가 너와 함께 함이라 놀라지 말라 나는 네 하나님이 됨이라 "
     "내가 너를 굳세게 하리라 참으로 너를 도와 주리라",
     "Isa 41:10", "So do not fear, for I am with you; do not be dismayed, for I am your God. "
     "I will strengthen you and help you.", "함께하심"),
    ("이사야 43:1", "너는 두려워하지 말라 내가 너를 구속하였고 내가 너를 지명하여 불렀나니 "
     "너는 내 것이라",
     "Isa 43:1", "Do not fear, for I have redeemed you; I have summoned you by name; you are mine.",
     "너는 내 것이라"),
    ("이사야 43:19", "보라 내가 새 일을 행하리니 이제 나타낼 것이라 너희가 그것을 알지 못하겠느냐",
     "Isa 43:19", "See, I am doing a new thing! Now it springs up; do you not perceive it?", "새 일"),
    ("이사야 55:8", "이는 내 생각이 너희의 생각과 다르며 내 길은 너희의 길과 다름이니라",
     "Isa 55:8", "“For my thoughts are not your thoughts, neither are your ways my ways,” "
     "declares the LORD.", "다른 생각 다른 길"),
    ("이사야 6:8", "내가 누구를 보내며 누가 우리를 위하여 갈꼬 그 때에 내가 이르되 "
     "내가 여기 있나이다 나를 보내소서",
     "Isa 6:8", "Then I heard the voice of the Lord saying, “Whom shall I send? And who will "
     "go for us?” And I said, “Here am I. Send me!”", "나를 보내소서"),
    ("이사야 9:6", "이는 한 아기가 우리에게 났고 한 아들을 우리에게 주신 바 되었는데 "
     "그의 어깨에는 정사를 메었고",
     "Isa 9:6", "For to us a child is born, to us a son is given, and the government will be on "
     "his shoulders.", "한 아기가 나셨네"),
    ("예레미야 29:11", "여호와의 말씀이니라 너희를 향한 나의 생각을 내가 아나니 평안이요 재앙이 "
     "아니니라 너희에게 미래와 희망을 주는 것이니라",
     "Jer 29:11", "“For I know the plans I have for you,” declares the LORD, “plans "
     "to prosper you and not to harm you, plans to give you hope and a future.”", "미래와 희망"),
    ("예레미야 33:3", "너는 내게 부르짖으라 내가 네게 응답하겠고 네가 알지 못하는 크고 은밀한 일을 "
     "네게 보이리라",
     "Jer 33:3", "Call to me and I will answer you and tell you great and unsearchable things "
     "you do not know.", "부르짖으라"),
    ("예레미야애가 3:22-23", "여호와의 인자와 긍휼이 무궁하시므로 우리가 진멸되지 아니함이니이다 "
     "이것들이 아침마다 새로우니 주의 성실하심이 크시도소이다",
     "La 3:22-23", "Because of the LORD's great love we are not consumed, for his compassions "
     "never fail. They are new every morning; great is your faithfulness.", "아침마다 새로우니"),
    ("미가 6:8", "사람아 주께서 선한 것이 무엇임을 네게 보이셨나니 여호와께서 네게 구하시는 것은 "
     "오직 정의를 행하며 인자를 사랑하며 겸손하게 네 하나님과 함께 행하는 것이 아니냐",
     "Mic 6:8", "He has shown you, O mortal, what is good. And what does the LORD require of you? "
     "To act justly and to love mercy and to walk humbly with your God.", "정의와 인자와 겸손"),
    ("하박국 3:18", "나는 여호와로 말미암아 즐거워하며 나의 구원의 하나님으로 말미암아 기뻐하리로다",
     "Hab 3:18", "Yet I will rejoice in the LORD, I will be joyful in God my Savior.", "그래도 기뻐하리"),
    ("스바냐 3:17", "너의 하나님 여호와가 너의 가운데에 계시니 그는 구원을 베푸실 전능자이시라 "
     "그가 너로 말미암아 기쁨을 이기지 못하시며",
     "Zep 3:17", "The LORD your God is with you, the Mighty Warrior who saves. He will take great "
     "delight in you.", "기쁨을 이기지 못하심"),
    ("여호수아 1:9", "내가 네게 명령한 것이 아니냐 강하고 담대하라 두려워하지 말며 놀라지 말라 "
     "네가 어디로 가든지 네 하나님 여호와가 너와 함께 하느니라",
     "Jos 1:9", "Have I not commanded you? Be strong and courageous. Do not be afraid; do not be "
     "discouraged, for the LORD your God will be with you wherever you go.", "강하고 담대하라"),
    ("신명기 31:6", "너희는 강하고 담대하라 두려워하지 말라 그들 앞에서 떨지 말라 이는 네 하나님 "
     "여호와 그가 너와 함께 가시며 결코 너를 떠나지 아니하시며 버리지 아니하실 것임이라",
     "Dt 31:6", "Be strong and courageous. Do not be afraid or terrified because of them, for the "
     "LORD your God goes with you; he will never leave you nor forsake you.", "떠나지 아니하심"),
    ("신명기 6:5", "너는 마음을 다하고 뜻을 다하고 힘을 다하여 네 하나님 여호와를 사랑하라",
     "Dt 6:5", "Love the LORD your God with all your heart and with all your soul and with all "
     "your strength.", "다하여 사랑하라"),
    ("출애굽기 14:14", "여호와께서 너희를 위하여 싸우시리니 너희는 가만히 있을지니라",
     "Ex 14:14", "The LORD will fight for you; you need only to be still.", "가만히 있을지니라"),
    ("사무엘상 16:7", "내가 보는 것은 사람과 같지 아니하니 사람은 외모를 보거니와 "
     "나 여호와는 중심을 보느니라",
     "1Sa 16:7", "The LORD does not look at the things people look at. People look at the outward "
     "appearance, but the LORD looks at the heart.", "중심을 보시는 분"),
    ("역대하 7:14", "내 이름으로 일컫는 내 백성이 그들의 악한 길에서 떠나 스스로 낮추고 기도하여 "
     "내 얼굴을 찾으면 내가 하늘에서 듣고 그들의 죄를 사하고 그들의 땅을 고칠지라",
     "2Ch 7:14", "If my people, who are called by my name, will humble themselves and pray and "
     "seek my face and turn from their wicked ways, then I will hear from heaven, and I will "
     "forgive their sin and will heal their land.", "땅을 고치시리라"),
    ("느헤미야 8:10", "여호와로 인하여 기뻐하는 것이 너희의 힘이니라",
     "Ne 8:10", "The joy of the LORD is your strength.", "기쁨이 힘이라"),
    ("욥기 23:10", "내가 가는 길을 그가 아시나니 그가 나를 단련하신 후에는 내가 순금 같이 나오리라",
     "Job 23:10", "But he knows the way that I take; when he has tested me, I will come forth as "
     "gold.", "순금 같이"),
    ("마태복음 5:16", "이같이 너희 빛이 사람 앞에 비치게 하여 그들로 너희 착한 행실을 보고 "
     "하늘에 계신 너희 아버지께 영광을 돌리게 하라",
     "Mt 5:16", "Let your light shine before others, that they may see your good deeds and "
     "glorify your Father in heaven.", "빛을 비추라"),
    ("마태복음 6:33", "그런즉 너희는 먼저 그의 나라와 그의 의를 구하라 그리하면 이 모든 것을 "
     "너희에게 더하시리라",
     "Mt 6:33", "But seek first his kingdom and his righteousness, and all these things will be "
     "given to you as well.", "먼저 그의 나라를"),
    ("마태복음 11:28", "수고하고 무거운 짐 진 자들아 다 내게로 오라 내가 너희를 쉬게 하리라",
     "Mt 11:28", "Come to me, all you who are weary and burdened, and I will give you rest.", "쉼"),
    ("마태복음 28:20", "내가 세상 끝날까지 너희와 항상 함께 있으리라",
     "Mt 28:20", "And surely I am with you always, to the very end of the age.", "항상 함께"),
    ("마가복음 10:45", "인자가 온 것은 섬김을 받으려 함이 아니라 도리어 섬기려 하고 "
     "자기 목숨을 많은 사람의 대속물로 주려 함이니라",
     "Mk 10:45", "For even the Son of Man did not come to be served, but to serve, and to give "
     "his life as a ransom for many.", "섬기러 오심"),
    ("누가복음 6:31", "남에게 대접을 받고자 하는 대로 너희도 남을 대접하라",
     "Lk 6:31", "Do to others as you would have them do to you.", "대접하라"),
    ("요한복음 3:16", "하나님이 세상을 이처럼 사랑하사 독생자를 주셨으니 이는 그를 믿는 자마다 "
     "멸망하지 않고 영생을 얻게 하려 하심이라",
     "Jn 3:16", "For God so loved the world that he gave his one and only Son, that whoever "
     "believes in him shall not perish but have eternal life.", "이처럼 사랑하사"),
    ("요한복음 14:6", "내가 곧 길이요 진리요 생명이니 나로 말미암지 않고는 아버지께로 올 자가 "
     "없느니라",
     "Jn 14:6", "I am the way and the truth and the life. No one comes to the Father except "
     "through me.", "길이요 진리요 생명"),
    ("요한복음 14:27", "평안을 너희에게 끼치노니 곧 나의 평안을 너희에게 주노라",
     "Jn 14:27", "Peace I leave with you; my peace I give you.", "나의 평안"),
    ("요한복음 15:5", "나는 포도나무요 너희는 가지라 그가 내 안에, 내가 그 안에 거하면 사람이 "
     "열매를 많이 맺나니 나를 떠나서는 너희가 아무 것도 할 수 없음이라",
     "Jn 15:5", "I am the vine; you are the branches. If you remain in me and I in you, you will "
     "bear much fruit; apart from me you can do nothing.", "가지로 붙어 있음"),
    ("사도행전 1:8", "오직 성령이 너희에게 임하시면 너희가 권능을 받고 예루살렘과 온 유대와 "
     "사마리아와 땅 끝까지 이르러 내 증인이 되리라",
     "Ac 1:8", "But you will receive power when the Holy Spirit comes on you; and you will be my "
     "witnesses in Jerusalem, and in all Judea and Samaria, and to the ends of the earth.",
     "땅 끝까지"),
    ("사도행전 20:35", "주 예수께서 친히 말씀하신 바 주는 것이 받는 것보다 복이 있다",
     "Ac 20:35", "Remembering the words the Lord Jesus himself said: “It is more blessed to "
     "give than to receive.”", "주는 것이 복"),
    ("로마서 5:8", "우리가 아직 죄인 되었을 때에 그리스도께서 우리를 위하여 죽으심으로 하나님께서 "
     "우리에 대한 자기의 사랑을 확증하셨느니라",
     "Ro 5:8", "But God demonstrates his own love for us in this: While we were still sinners, "
     "Christ died for us.", "확증된 사랑"),
    ("로마서 8:28", "우리가 알거니와 하나님을 사랑하는 자 곧 그의 뜻대로 부르심을 입은 자들에게는 "
     "모든 것이 합력하여 선을 이루느니라",
     "Ro 8:28", "And we know that in all things God works for the good of those who love him, who "
     "have been called according to his purpose.", "합력하여 선을"),
    ("로마서 12:2", "너희는 이 세대를 본받지 말고 오직 마음을 새롭게 함으로 변화를 받아",
     "Ro 12:2", "Do not conform to the pattern of this world, but be transformed by the renewing "
     "of your mind.", "변화를 받아"),
    ("로마서 12:15", "즐거워하는 자들과 함께 즐거워하고 우는 자들과 함께 울라",
     "Ro 12:15", "Rejoice with those who rejoice; mourn with those who mourn.", "함께 울라"),
    ("고린도전서 13:13", "그런즉 믿음, 소망, 사랑, 이 세 가지는 항상 있을 것인데 "
     "그 중의 제일은 사랑이라",
     "1Co 13:13", "And now these three remain: faith, hope and love. But the greatest of these "
     "is love.", "제일은 사랑"),
    ("고린도전서 10:13", "사람이 감당할 시험 밖에는 너희가 당한 것이 없나니 오직 하나님은 미쁘사 "
     "너희가 감당하지 못할 시험 당함을 허락하지 아니하시고",
     "1Co 10:13", "No temptation has overtaken you except what is common to mankind. And God is "
     "faithful; he will not let you be tempted beyond what you can bear.", "감당할 시험"),
    ("고린도후서 5:17", "그런즉 누구든지 그리스도 안에 있으면 새로운 피조물이라 이전 것은 "
     "지나갔으니 보라 새 것이 되었도다",
     "2Co 5:17", "Therefore, if anyone is in Christ, the new creation has come: The old has gone, "
     "the new is here!", "새로운 피조물"),
    ("고린도후서 12:9", "내 은혜가 네게 족하도다 이는 내 능력이 약한 데서 온전하여짐이라",
     "2Co 12:9", "My grace is sufficient for you, for my power is made perfect in weakness.",
     "약한 데서 온전하여짐"),
    ("갈라디아서 2:20", "내가 그리스도와 함께 십자가에 못 박혔나니 그런즉 이제는 내가 사는 것이 "
     "아니요 오직 내 안에 그리스도께서 사시는 것이라",
     "Gal 2:20", "I have been crucified with Christ and I no longer live, but Christ lives in me.",
     "함께 못 박힘"),
    ("갈라디아서 6:9", "우리가 선을 행하되 낙심하지 말지니 포기하지 아니하면 때가 이르매 거두리라",
     "Gal 6:9", "Let us not become weary in doing good, for at the proper time we will reap a "
     "harvest if we do not give up.", "낙심하지 말지니"),
    ("에베소서 2:8", "너희는 그 은혜에 의하여 믿음으로 말미암아 구원을 받았으니 이것은 너희에게서 "
     "난 것이 아니요 하나님의 선물이라",
     "Eph 2:8", "For it is by grace you have been saved, through faith—and this is not from "
     "yourselves, it is the gift of God.", "은혜로 구원"),
    ("에베소서 4:32", "서로 친절하게 하며 불쌍히 여기며 서로 용서하기를 하나님이 그리스도 안에서 "
     "너희를 용서하심과 같이 하라",
     "Eph 4:32", "Be kind and compassionate to one another, forgiving each other, just as in "
     "Christ God forgave you.", "서로 용서하라"),
    ("빌립보서 4:6", "아무것도 염려하지 말고 다만 모든 일에 기도와 간구로 너희 구할 것을 "
     "감사함으로 하나님께 아뢰라",
     "Php 4:6", "Do not be anxious about anything, but in every situation, by prayer and "
     "petition, with thanksgiving, present your requests to God.", "염려를 기도로"),
    ("빌립보서 4:13", "내게 능력 주시는 자 안에서 내가 모든 것을 할 수 있느니라",
     "Php 4:13", "I can do all this through him who gives me strength.", "능력 주시는 자"),
    ("골로새서 3:23", "무슨 일을 하든지 마음을 다하여 주께 하듯 하고 사람에게 하듯 하지 말라",
     "Col 3:23", "Whatever you do, work at it with all your heart, as working for the Lord, not "
     "for human masters.", "주께 하듯"),
    ("데살로니가전서 5:16-18", "항상 기뻐하라 쉬지 말고 기도하라 범사에 감사하라 이것이 그리스도 "
     "예수 안에서 너희를 향하신 하나님의 뜻이니라",
     "1Th 5:16-18", "Rejoice always, pray continually, give thanks in all circumstances; for this "
     "is God's will for you in Christ Jesus.", "기뻐하고 기도하고 감사하라"),
    ("디모데후서 1:7", "하나님이 우리에게 주신 것은 두려워하는 마음이 아니요 오직 능력과 사랑과 "
     "절제하는 마음이니",
     "2Ti 1:7", "For the Spirit God gave us does not make us timid, but gives us power, love and "
     "self-discipline.", "두려움이 아니라"),
    ("히브리서 11:1", "믿음은 바라는 것들의 실상이요 보이지 않는 것들의 증거니",
     "Heb 11:1", "Now faith is confidence in what we hope for and assurance about what we do not "
     "see.", "믿음의 실상"),
    ("히브리서 12:1", "이러므로 우리에게 구름 같이 둘러싼 허다한 증인들이 있으니 모든 무거운 것과 "
     "얽매이기 쉬운 죄를 벗어 버리고 인내로써 우리 앞에 당한 경주를 하며",
     "Heb 12:1", "Therefore, since we are surrounded by such a great cloud of witnesses, let us "
     "throw off everything that hinders and the sin that so easily entangles. And let us run with "
     "perseverance the race marked out for us.", "믿음의 경주"),
    ("히브리서 13:8", "예수 그리스도는 어제나 오늘이나 영원토록 동일하시니라",
     "Heb 13:8", "Jesus Christ is the same yesterday and today and forever.", "동일하신 주"),
    ("야고보서 1:2-3", "내 형제들아 너희가 여러 가지 시험을 당하거든 온전히 기쁘게 여기라 이는 "
     "너희 믿음의 시련이 인내를 만들어 내는 줄 너희가 앎이라",
     "Jas 1:2-3", "Consider it pure joy, my brothers and sisters, whenever you face trials of many "
     "kinds, because you know that the testing of your faith produces perseverance.", "시련과 인내"),
    ("베드로전서 5:7", "너희 염려를 다 주께 맡기라 이는 그가 너희를 돌보심이라",
     "1Pe 5:7", "Cast all your anxiety on him because he cares for you.", "돌보심"),
    ("요한일서 1:9", "만일 우리가 우리 죄를 자백하면 그는 미쁘시고 의로우사 우리 죄를 사하시며 "
     "우리를 모든 불의에서 깨끗하게 하실 것이요",
     "1Jn 1:9", "If we confess our sins, he is faithful and just and will forgive us our sins and "
     "purify us from all unrighteousness.", "자백하면"),
    ("요한일서 4:19", "우리가 사랑함은 그가 먼저 우리를 사랑하셨음이라",
     "1Jn 4:19", "We love because he first loved us.", "먼저 사랑하심"),
    ("요한계시록 3:20", "볼지어다 내가 문 밖에 서서 두드리노니 누구든지 내 음성을 듣고 문을 열면 "
     "내가 그에게로 들어가 그와 더불어 먹고 그는 나와 더불어 먹으리라",
     "Rev 3:20", "Here I am! I stand at the door and knock. If anyone hears my voice and opens "
     "the door, I will come in and eat with that person, and they with me.", "문 밖에서 두드리심"),
]

# ==============================================================================
# 예화 71편 — 실제 역사·인물·성경 본문·작품 창작 배경만
# ==============================================================================
STORIES = [
    ("내 영혼 평안해",
     "1873년 미국의 변호사 호레이쇼 스패퍼드는 먼저 유럽으로 떠난 아내와 네 딸을 뒤따라가던 "
     "길이었다. 딸들이 탄 여객선 빌 뒤 아브르호가 대서양에서 충돌해 침몰했고, 아내만 살아남아 "
     "'나 홀로 구조되었음'이라는 전보를 보냈다. 그는 딸들이 가라앉은 바다 위를 지나며 찬송시 "
     "「내 평생에 가는 길」을 썼다."),
    ("라벤스브뤼크의 자매",
     "네덜란드 시계공 코리 텐 붐과 언니 벳시는 유대인을 집에 숨겨 준 일로 1944년 라벤스브뤼크 "
     "수용소에 갇혔다. 몰래 들여온 성경을 함께 읽으며 두 사람은 밤마다 동료 수감자들과 예배를 "
     "드렸다. 벳시는 그곳에서 숨을 거두었고, 코리는 행정 착오로 풀려나 평생 용서를 전하며 살았다."),
    ("노예선 선장의 노래",
     "존 뉴턴은 18세기 영국의 노예무역선 선장이었다. 1748년 폭풍우 속에서 가까스로 살아난 일을 "
     "계기로 신앙을 되찾았고, 뒤늦게 목사가 되어 1772년 「나 같은 죄인 살리신」을 썼다. 말년에는 "
     "노예무역 폐지 운동에 가담해 윌리엄 윌버포스를 도왔다."),
    ("스물넉 달이 아니라 스물나흘",
     "1741년 여름, 게오르크 프리드리히 헨델은 빚에 몰리고 건강이 무너져 재기가 어렵다는 말을 "
     "듣던 처지였다. 그해 8월부터 9월까지 방에 틀어박혀 스물네 날 만에 오라토리오 「메시아」를 "
     "완성했다. 이듬해 더블린 초연의 수익은 감옥에 갇힌 채무자들을 풀어 주는 데 쓰였다."),
    ("비텐베르크의 성문",
     "1517년 10월 31일, 마르틴 루터는 비텐베르크 성교회 문에 95개조 반박문을 붙였다. 뒷날 "
     "바르트부르크 성에 숨어 지내던 열 달 동안 그는 신약성경을 독일어로 옮겼다. 라틴어를 모르는 "
     "사람도 성경을 직접 읽게 되자, 발밑을 비추는 등불이 소수의 손에서 모두의 손으로 옮겨 갔다."),
    ("고아 이천 명의 아침 식탁",
     "조지 뮐러는 젊은 시절 도둑질과 사기로 감옥에 갇혔던 사람이었다. 회심한 뒤 1836년 영국 "
     "브리스톨에서 고아원을 시작해 평생 만 명이 넘는 아이들을 돌보았다. 그는 후원을 요청하는 "
     "편지를 한 번도 보내지 않고 기도로만 운영했다고 기록에 남겼다."),
    ("굶주림의 막사에서",
     "1941년 아우슈비츠에서 탈주자가 발생하자 수용소는 본보기로 열 사람을 굶겨 죽이기로 했다. "
     "뽑힌 사람 가운데 하나가 가족을 부르짖자, 폴란드인 신부 막시밀리안 콜베가 앞으로 나와 그를 "
     "대신하겠다고 했다. 그가 대신한 프란치셰크 가요브니체크는 살아남아 전쟁 뒤 오래 살았다."),
    ("물맷돌 다섯 개",
     "블레셋의 골리앗 앞에서 사울은 다윗에게 자기 갑주를 입혔다. 다윗은 익숙하지 않다며 그것을 "
     "벗고, 시내에서 매끄러운 돌 다섯 개를 골라 목자의 주머니에 넣었다(삼상 17:38-40). 그가 "
     "의지한 것은 무기의 크기가 아니라 자기를 사자와 곰에게서 건지신 분이었다."),
    ("사자 굴의 창문",
     "다니엘은 왕의 금령이 내려진 것을 알고도 하루 세 번 예루살렘을 향한 창문을 열고 기도했다"
     "(단 6:10). 그는 금령을 몰라서가 아니라 알고도 늘 하던 대로 했다. 위기 앞에서 새로 시작한 "
     "경건이 아니라, 평소에 몸에 밴 경건이 그를 지켰다."),
    ("애굽의 총리가 된 노예",
     "요셉은 형들에게 팔려 애굽으로 끌려갔고, 억울한 누명으로 감옥에 갇혔다. 뒷날 형들 앞에 선 "
     "그는 '당신들은 나를 해하려 하였으나 하나님은 그것을 선으로 바꾸사'라고 말했다(창 50:20). "
     "긴 세월이 지나서야 보이는 뜻이 있었다."),
    ("모라비아의 백 년 기도",
     "1727년 8월 13일 독일 헤른후트의 모라비아 공동체에 큰 부흥이 일어났다. 이들은 24시간 "
     "끊이지 않는 기도를 시작했고, 그 기도가 백 년 넘게 이어졌다고 전한다. 이 작은 마을에서 "
     "수백 명의 선교사가 세계 각지로 나갔다."),
    ("올더스게이트의 저녁",
     "1738년 5월 24일 저녁, 존 웨슬리는 런던 올더스게이트 거리의 한 모임에 마지못해 참석했다. "
     "루터의 로마서 서문이 낭독되던 중 그는 '내 마음이 이상하게 뜨거워짐을 느꼈다'고 일기에 "
     "적었다. 선교사로 실패하고 돌아온 뒤의 일이었다."),
    ("석탄 광부의 아이들",
     "18세기 영국의 조지 휫필드와 존 웨슬리는 교회 문이 닫히자 브리스톨 킹스우드의 탄광 앞 "
     "들판에서 설교했다. 얼굴이 검댕으로 덮인 광부들의 뺨에 눈물이 흘러 흰 줄이 생겼다고 "
     "휫필드는 기록했다. 강단이 없어서 강단이 넓어진 자리였다."),
    ("아프리카의 지도 밖으로",
     "데이비드 리빙스턴은 1841년 아프리카에 도착해 삼십 년간 대륙을 걸었다. 그는 유럽 지도에 "
     "비어 있던 내륙을 기록했고, 가는 곳마다 노예무역의 참상을 세상에 알렸다. 1873년 잠비아의 "
     "한 마을에서 무릎을 꿇은 자세로 숨진 채 발견되었다."),
    ("중국 내지의 옷",
     "허드슨 테일러는 1854년 중국에 도착한 뒤, 서양 선교사들과 달리 변발을 하고 중국옷을 입었다. "
     "동료들의 비웃음을 샀지만 내륙 사람들은 그를 받아들였다. 그가 세운 중국내지선교회는 그가 "
     "세상을 떠날 무렵 팔백 명이 넘는 선교사를 파송하고 있었다."),
    ("양화진의 스물다섯",
     "1885년 조선에 온 미국인 선교사 헨리 아펜젤러는 배재학당을 세웠고, 1902년 목포로 가는 배가 "
     "충돌해 침몰할 때 함께 탄 한국인 학생을 구하려다 숨졌다. 서울 양화진에는 그를 비롯해 조선 "
     "땅에 묻히기를 택한 선교사들의 무덤이 남아 있다."),
    ("에콰도르의 다섯 청년",
     "1956년 1월, 짐 엘리엇을 비롯한 다섯 명의 미국인 선교사가 에콰도르 아우카족에게 복음을 "
     "전하러 갔다가 창에 찔려 죽었다. 몇 해 뒤 엘리엇의 아내 엘리자베스와 동료의 누이가 그 "
     "부족에게 들어가 함께 살았고, 창을 들었던 사람들 가운데 신자가 나왔다."),
    ("성경을 밭에 묻고",
     "일제강점기와 6·25 전쟁을 지나며 많은 한국 교회 성도들이 성경을 지키려 애썼다. 신사참배를 "
     "거부하고 옥에 갇힌 주기철 목사는 1944년 평양형무소에서 숨졌다. 그의 마지막 설교 제목은 "
     "'다섯 종목의 나의 기원'이었다."),
    ("한 손에 성경 한 손에 삽",
     "1907년 평양 장대현교회에서 시작된 사경회에서 길선주 목사의 통회를 시작으로 대부흥이 "
     "일어났다. 사람들은 밤새 자기 죄를 자복했고, 훔친 물건을 돌려주고 원수와 화해하는 일이 "
     "잇따랐다. 회개가 예배당 밖의 삶을 바꾸어 놓았다."),
    ("스코틀랜드의 달리는 선교사",
     "1924년 파리 올림픽에서 에릭 리델은 주 종목인 100m 예선이 주일에 열린다는 이유로 출전을 "
     "포기했다. 대신 훈련도 부족했던 400m에 나가 세계신기록으로 우승했다. 그는 이후 중국 "
     "선교사로 갔고 1945년 일본군 수용소에서 숨졌다."),
    ("휠체어에 앉은 화가",
     "조니 에릭슨 타다는 1967년 열일곱 살에 다이빙 사고로 목 아래가 마비되었다. 몇 해의 절망을 "
     "지난 뒤 그는 입에 붓을 물고 그림을 그리기 시작했고, 이후 장애인 사역 단체를 세워 세계 여러 "
     "나라에 휠체어를 보내는 일을 해 왔다."),
    ("버밍햄 감옥의 편지",
     "1963년 4월, 마틴 루터 킹은 앨라배마 버밍햄 감옥에 갇혀 신문 여백과 화장지에 편지를 썼다. "
     "'기다리라'는 성직자들의 충고에 답한 이 글은 뒷날 「버밍햄 감옥에서 보내는 편지」로 알려졌다."),
    ("의회의 스무 해",
     "윌리엄 윌버포스는 1787년부터 영국 의회에서 노예무역 폐지 법안을 거듭 발의했다. 거듭 "
     "부결되었지만 그는 스무 해를 물러서지 않았고, 1807년 마침내 노예무역 폐지법이 통과되었다. "
     "노예제 자체를 폐지하는 법은 그가 숨지기 사흘 전인 1833년에 통과되었다."),
    ("등불을 든 여인",
     "플로렌스 나이팅게일은 1854년 크림 전쟁 때 서른여덟 명의 간호사를 이끌고 스쿠타리 야전병원에 "
     "갔다. 병원의 위생을 뜯어고쳐 사망률을 크게 낮추었고, 밤마다 등불을 들고 병상을 돌았다. "
     "그는 열일곱 살에 하나님의 부르심을 들었다고 일기에 적었다."),
    ("빈민가의 구세군",
     "윌리엄 부스는 1865년 런던 이스트엔드의 빈민가에서 전도를 시작했다. 술과 굶주림에 무너진 "
     "사람들에게 설교만 하지 않고 수프와 잠자리를 내주었다. 그가 세운 구세군은 '수프, 비누, "
     "구원'이라는 말로 그 방식을 요약했다."),
    ("암스테르담의 뒤뜰",
     "네덜란드의 안네 프랑크 가족은 1942년부터 암스테르담의 한 건물 뒤편 은신처에 숨어 지냈다. "
     "회사 직원 미프 히스를 비롯한 몇 사람이 이 년 넘게 몰래 음식을 날랐다. 뒷날 미프는 "
     "'나는 영웅이 아니다. 그저 해야 할 일을 했을 뿐이다'라고 말했다."),
    ("르 샹봉 마을",
     "제2차 세계대전 중 프랑스 산간 마을 르 샹봉쉬르리뇽의 주민들은 앙드레 트로크메 목사의 "
     "인도 아래 유대인 수천 명을 숨겨 주었다. 위그노의 후손인 이들은 자기 조상이 박해받던 기억을 "
     "잊지 않았다. 마을 전체가 한 사람도 밀고하지 않았다."),
    ("스기하라의 비자",
     "1940년 리투아니아 주재 일본 영사 스기하라 지우네는 본국의 반대를 무릅쓰고 유대인 난민들에게 "
     "통과 비자를 써 주었다. 영사관이 폐쇄되어 기차에 오르는 순간까지 손으로 비자를 써서 창밖으로 "
     "건넸다. 그렇게 살아난 사람이 육천 명에 이른다고 전한다."),
    ("쉰들러의 명단",
     "독일인 사업가 오스카 쉰들러는 자기 공장에 '필수 인력'이라는 명목으로 유대인들을 등록해 "
     "천이백 명가량을 수용소에서 빼냈다. 전쟁이 끝났을 때 그의 재산은 남아 있지 않았다. "
     "그는 '한 사람을 구하는 자는 온 세계를 구하는 것'이라는 말을 들었다."),
    ("남극의 스물여덟",
     "1914년 어니스트 섀클턴의 남극 탐험선 인듀어런스호가 얼음에 갇혀 부서졌다. 그는 대원 "
     "스물일곱 명을 이끌고 유빙 위에서 이 년 가까이 버텼고, 구명보트로 사우스조지아섬까지 건너가 "
     "구조를 데려왔다. 한 사람도 잃지 않았다."),
    ("아폴로 13호",
     "1970년 4월, 달로 향하던 아폴로 13호의 산소 탱크가 폭발했다. 휴스턴 관제실은 우주선에 실린 "
     "물건만으로 이산화탄소 제거 장치를 다시 만들어 무전으로 조립법을 불러 주었다. 세 사람은 "
     "나흘 뒤 태평양에 무사히 내려왔다."),
    ("칠레 광산의 예순아홉 날",
     "2010년 8월 칠레 산호세 광산이 무너져 광부 서른세 명이 지하 700m에 갇혔다. 열이레 만에 "
     "생존이 확인되었고, 예순아홉 날 만에 전원이 구조되었다. 갇힌 동안 이들은 조를 나눠 매일 "
     "기도 모임을 열었다고 증언했다."),
    ("허드슨강의 착륙",
     "2009년 1월, 뉴욕 라과디아 공항을 이륙한 여객기가 새 떼와 부딪혀 두 엔진을 모두 잃었다. "
     "기장 체슬리 설런버거는 삼 분 만에 판단해 허드슨강에 비상 착수했고, 승객과 승무원 백오십오 "
     "명 전원이 살아남았다. 그는 마지막으로 기내를 두 번 걸어 확인한 뒤 내렸다."),
    ("한 사람의 서명",
     "1983년 9월, 소련 방공군 중령 스타니슬라프 페트로프는 미국이 핵미사일을 발사했다는 경보를 "
     "받았다. 그는 규정대로 상부에 보고하는 대신 오작동이라고 판단했고, 실제로 위성의 오류였다. "
     "그의 한 번의 판단이 전쟁을 막았다."),
    ("전쟁터의 크리스마스",
     "1914년 12월 24일 서부전선의 참호에서 독일군이 부른 「고요한 밤」을 듣고 영국군이 화답했다. "
     "이튿날 병사들은 무인지대로 걸어 나와 악수하고 함께 전사자를 묻었다. 지휘부는 이 일이 "
     "되풀이되지 않도록 이듬해부터 엄히 단속했다."),
    ("사막의 우물",
     "하갈이 브엘세바 광야에서 물이 떨어지자 아이를 덤불 아래 두고 멀찍이 앉아 울었다. 하나님이 "
     "그의 눈을 밝히시니 샘물이 보였다(창 21:19). 우물은 없다가 생긴 것이 아니라, 보이지 않던 "
     "것이 보인 것이었다."),
    ("갈멜산의 열두 통",
     "엘리야는 바알 선지자들과의 대결에서 제단 위에 물 열두 통을 부어 도랑까지 채웠다"
     "(왕상 18:33-35). 불이 붙기 가장 어렵게 만든 뒤에 하나님께 구했다. 그리고 자기 힘으로 "
     "설명될 여지를 남기지 않았다."),
    ("로뎀나무 아래",
     "갈멜산의 승리 직후 엘리야는 이세벨의 위협에 쫓겨 광야로 도망쳐 죽기를 구했다. 하나님은 "
     "책망 대신 천사를 보내 떡과 물을 주고 다시 자게 하셨다(왕상 19:5-7). 무너진 사람에게 먼저 "
     "필요한 것은 설교가 아니라 빵과 잠이었다."),
    ("과부의 두 렙돈",
     "예수께서 헌금함 맞은편에 앉아 사람들이 돈 넣는 것을 보고 계셨다. 한 가난한 과부가 두 렙돈을 "
     "넣자 제자들을 불러 그가 가장 많이 넣었다고 하셨다(막 12:41-44). 금액이 아니라 남은 것을 "
     "보신 것이다."),
    ("보리떡 다섯 개",
     "빈 들에서 예수께서는 한 아이가 가진 보리떡 다섯 개와 물고기 두 마리를 받으셨다"
     "(요 6:9). 안드레조차 '이것이 얼마나 되겠사옵나이까'라고 했지만, 예수는 그 작은 것을 "
     "손에 드시고 축사하셨다."),
    ("지붕을 뜯은 친구들",
     "가버나움의 한 집에 사람이 가득해 들어갈 수 없자, 네 친구가 중풍병자를 침상째 들고 지붕에 "
     "올라가 흙을 파헤치고 달아 내렸다(막 2:4). 예수께서 보신 것은 병자의 믿음만이 아니라 "
     "'그들의 믿음'이었다."),
    ("돌아온 아들의 신발",
     "아버지는 멀리서 아들을 보고 달려가 목을 안고 입을 맞추었다. 그리고 종들에게 제일 좋은 옷과 "
     "가락지와 신을 가져오라 했다(눅 15:22). 신을 신긴 것은 그가 종이 아니라 아들로 돌아왔음을 "
     "집안 모두에게 알리는 일이었다."),
    ("사마리아 사람의 두 데나리온",
     "강도 만난 사람을 보고 제사장과 레위인은 피해 지나갔다. 사마리아 사람은 상처를 싸매고 자기 "
     "짐승에 태워 주막으로 데려간 뒤, 두 데나리온을 주며 '비용이 더 들면 돌아올 때 갚겠다'고 "
     "했다(눅 10:35). 그는 한 번의 선행으로 끝내지 않았다."),
    ("새벽 세 시의 파도",
     "제자들이 밤새 노를 저어도 배가 나아가지 않던 새벽, 예수께서 바다 위로 걸어오셨다"
     "(막 6:48). 성경은 그가 '그들이 괴로이 노 젓는 것을 보시고' 오셨다고 적는다. 보고 계셨다는 "
     "말이 먼저 나온다."),
    ("닭 울기 전에",
     "베드로는 세 번 부인한 뒤 닭 우는 소리를 듣고 밖에 나가 심히 통곡했다. 부활 후 디베랴 "
     "바닷가에서 예수는 그에게 세 번 '네가 나를 사랑하느냐'고 물으셨다(요 21:15-17). 세 번의 "
     "부인에 세 번의 회복이 있었다."),
    ("다메섹 도상",
     "사울은 그리스도인을 잡으러 다메섹으로 가던 길에 빛을 만나 사흘간 보지 못했다. 아나니아는 "
     "그를 두려워했지만 찾아가 '형제 사울아' 하고 불렀다(행 9:17). 교회가 가장 두려워하던 사람을 "
     "형제라 부른 첫 사람이었다."),
    ("빌립보 감옥의 찬송",
     "매를 맞고 착고에 채인 바울과 실라는 한밤중에 기도하며 하나님을 찬송했고 죄수들이 듣고 "
     "있었다(행 16:25). 지진으로 문이 열렸을 때 그들은 도망하지 않았고, 그 밤에 간수의 온 집이 "
     "믿게 되었다."),
    ("유두고의 창턱",
     "바울이 밤늦도록 강론할 때 유두고라는 청년이 창에 걸터앉아 졸다가 삼층에서 떨어져 죽었다"
     "(행 20:9). 바울은 내려가 그를 안고 살렸고, 사람들은 다시 올라가 날이 새기까지 이야기를 "
     "나누었다."),
    ("바구니에 담긴 사람",
     "다메섹에서 죽이려는 자들을 피해 제자들이 밤에 바울을 광주리에 담아 성벽에서 달아 내렸다"
     "(행 9:25). 훗날 그는 이 일을 자기 약함을 자랑하는 목록의 맨 끝에 적어 두었다"
     "(고후 11:33)."),
    ("바다 위의 밀 자루",
     "로마로 가던 배가 유라굴로 광풍에 열나흘을 떠돌 때, 바울은 이백칠십육 명 앞에서 떡을 떼어 "
     "감사하고 먹었다. 사람들이 안심하고 먹은 뒤 남은 밀을 바다에 버려 배를 가볍게 했다"
     "(행 27:35-38)."),
    ("느헤미야의 오십이 일",
     "느헤미야는 무너진 예루살렘 성벽을 오십이 일 만에 다시 쌓았다. 비결은 각 사람이 '자기 집 "
     "맞은편'을 맡은 것이었다(느 3장). 모두가 전체를 맡으려 하지 않고 자기 앞을 맡았다."),
    ("에스더의 사흘",
     "에스더는 왕 앞에 나아가기 전 수산에 있는 유다인들에게 사흘 금식을 청하고 '죽으면 죽으리이다' "
     "라고 했다(에 4:16). 그는 두려움이 없어서가 아니라 두려움을 안고 갔다."),
    ("룻의 이삭",
     "모압 여인 룻은 시어머니를 따라 낯선 땅에 와 보아스의 밭에서 이삭을 주웠다(룻 2장). 그는 "
     "아침부터 저녁까지 쉬지 않았고, 보아스는 일부러 곡식을 뽑아 흘려 두라고 일꾼들에게 일렀다."),
    ("사르밧 과부의 마지막 식사",
     "가뭄에 통의 가루 한 움큼과 병의 기름 조금이 전부였던 과부에게 엘리야는 먼저 자기에게 떡을 "
     "만들어 달라고 했다(왕상 17:12-13). 그가 먼저 내어놓자 통의 가루와 병의 기름이 떨어지지 "
     "않았다."),
    ("나아만의 일곱 번",
     "아람의 장군 나아만은 요단강에 일곱 번 몸을 씻으라는 말에 화를 냈다. 종들이 '큰일을 "
     "행하라 하였다면 하지 않았겠느냐'고 권하자 그제야 내려갔고 살이 어린아이 같아졌다"
     "(왕하 5:13-14)."),
    ("여리고의 일곱째 날",
     "이스라엘은 엿새 동안 하루 한 번씩 여리고를 잠잠히 돌았다. 일곱째 날에야 일곱 번 돌고 "
     "소리를 질렀다(수 6장). 아무 일도 일어나지 않는 엿새를 지나야 이레째가 있었다."),
    ("기드온의 삼백 명",
     "삼만 이천 명으로 시작한 기드온의 군대는 두려운 자를 돌려보내고 물 마시는 모습으로 걸러 "
     "삼백 명이 남았다(삿 7장). 하나님은 '너희 손이 나를 구원하였다 할까 하노라'고 이유를 "
     "밝히셨다."),
    ("한나의 서원",
     "한나는 성전에서 입술만 움직이며 기도해 엘리 제사장에게 취한 여자로 오해받았다"
     "(삼상 1:13). 그는 오해를 받으면서도 마음을 하나님 앞에 쏟아 놓기를 그치지 않았다."),
    ("솔로몬의 한 가지 구함",
     "기브온에서 하나님이 무엇을 줄까 물으시자 솔로몬은 부와 장수 대신 '듣는 마음'을 구했다"
     "(왕상 3:9). 그가 구하지 않은 것까지 함께 받은 이유가 거기 있었다."),
    ("바디매오의 겉옷",
     "여리고 길가의 맹인 바디매오는 예수께서 부르신다는 말을 듣고 겉옷을 내버리고 뛰어 일어나 "
     "나아갔다(막 10:50). 구걸할 때 깔고 앉던 그 옷을 버린 것이 그의 대답이었다."),
    ("삭개오의 나무",
     "키가 작아 볼 수 없던 세리장 삭개오는 앞질러 달려가 돌무화과나무에 올랐다. 예수는 그 아래에 "
     "이르러 이름을 부르셨다(눅 19:5). 군중이 죄인이라 수군거리던 이름을 먼저 부르신 것이다."),
    ("향유 옥합",
     "한 여자가 값진 향유 한 옥합을 깨뜨려 예수의 머리에 부었다. 사람들이 낭비라고 책망하자 "
     "예수는 '그가 내게 좋은 일을 하였느니라' 하셨다(막 14:6). 계산되지 않는 사랑이 있었다."),
    ("도마의 손가락",
     "도마는 못 자국에 손을 넣어 보지 않고는 믿지 않겠다고 했다. 여드레 뒤 예수는 그를 위해 다시 "
     "오셔서 손을 내미셨다(요 20:27). 의심하는 한 사람을 위해 따로 오신 걸음이었다."),
    ("엠마오의 저녁",
     "낙심해 고향으로 돌아가던 두 사람 곁에 예수께서 함께 걸으셨다. 그들은 떡을 떼실 때에야 눈이 "
     "밝아졌고, 곧바로 그 밤에 예루살렘으로 되돌아갔다(눅 24:31-33)."),
    ("사도의 겉옷과 가죽 종이",
     "바울은 마지막 편지에서 드로아에 두고 온 겉옷과 책, 특별히 가죽 종이를 가져오라고 부탁했다"
     "(딤후 4:13). 순교를 앞둔 사람의 부탁 목록에 여전히 읽을 것이 있었다."),
    ("옥중에서 쓴 편지",
     "존 번연은 허가 없이 설교했다는 이유로 열두 해를 베드퍼드 감옥에서 보냈다. 그곳에서 딸이 "
     "가져다준 실을 꼬아 구두끈을 만들어 팔며 「천로역정」을 썼다. 이 책은 이후 이백 개가 넘는 "
     "언어로 옮겨졌다."),
    ("점자와 열다섯 살",
     "루이 브라유는 세 살에 눈을 다쳐 시력을 잃었다. 열다섯 살이던 1824년, 군용 야간 암호에서 "
     "착안해 여섯 점으로 글자를 표현하는 방식을 만들었다. 그가 살아 있는 동안에는 인정받지 "
     "못했고, 사후에야 세계가 그 글자를 썼다."),
    ("귀먹은 작곡가의 초연",
     "1824년 빈에서 베토벤의 교향곡 9번이 초연되었다. 이미 소리를 듣지 못하던 그는 등을 돌린 채 "
     "박자를 젓고 있었고, 연주가 끝난 것도 몰랐다. 한 성악가가 그의 소매를 당겨 돌려세우자 "
     "청중이 일어나 손수건을 흔들고 있었다."),
    ("램프를 든 목사의 딸",
     "영국의 사라 트리머, 로버트 레이크스 같은 이들이 18세기 후반 주일학교를 시작했다. 주중에 "
     "공장에서 일하던 아이들이 주일에 모여 글을 배웠다. 성경을 읽히려고 시작한 교실이 그 나라의 "
     "문해율을 바꾸어 놓았다."),
    ("한 장의 지도",
     "1854년 런던 소호에 콜레라가 돌자 의사 존 스노는 사망자의 집을 지도에 하나씩 찍었다. 점들이 "
     "브로드가의 한 펌프를 둘러싸고 있었다. 그가 펌프 손잡이를 떼어 내자 유행이 잦아들었다. "
     "막연한 공포를 이긴 것은 성실한 기록이었다."),
    ("실패한 접착제",
     "1968년 3M의 연구원 스펜서 실버는 잘 붙지 않는 접착제를 만들어 놓고 쓸 데를 찾지 못했다. "
     "몇 해 뒤 동료 아트 프라이가 찬양대에서 찬송가 갈피가 자꾸 빠지는 일로 애를 먹다 그것을 "
     "떠올렸다. 포스트잇은 그렇게 나왔다."),
    ("소록도의 두 간호사",
     "오스트리아에서 온 마리안느 스퇴거와 마가레트 피사렉은 1960년대부터 사십 년 넘게 소록도에서 "
     "한센인들을 돌보았다. 장갑 없이 상처를 만졌고, 표창을 받는 자리에는 나가지 않았다. "
     "2005년 나이가 들어 짐이 될까 염려해 편지 한 장을 남기고 조용히 떠났다."),
]


def pick(day_of_year: int):
    """
    그날의 (구절, 예화)를 고른다.
    73 과 71 이 서로소이므로 1년 365일이 모두 다른 조합이 된다.
    """
    d = int(day_of_year)
    ref, ko, en_ref, en, topic = VERSES[d % len(VERSES)]
    title, body = STORIES[d % len(STORIES)]
    return {
        "topic": topic,
        "ko_ref": ref, "ko_verse": ko,
        "en_ref": en_ref, "en_verse": en,
        "story_title": title, "story_body": body,
    }


def stats():
    return {"verses": len(VERSES), "stories": len(STORIES),
            "unique_days": len(VERSES) * len(STORIES)}
