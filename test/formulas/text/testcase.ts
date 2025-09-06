import FormulaError from "../../../formulas/error";
export default {

    ASC: {
        'ASC("ＡＢＣ")': "ABC",
        'ASC("ヲァィゥ")': 'ｦｧｨｩ',
        'ASC("，。")': ',｡',
    },

    BAHTTEXT: {
        'BAHTTEXT(63147.89)': 'หกหมื่นสามพันหนึ่งร้อยสี่สิบเจ็ดบาทแปดสิบเก้าสตางค์',
        'BAHTTEXT(1234)': 'หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน',
    },

    CHAR: {
        'CHAR(65)': 'A',
        'CHAR(33)': '!',
    },

    CLEAN: {
        'CLEAN("äÄçÇéÉêPHP-MySQLöÖÐþúÚ")': "äÄçÇéÉêPHP-MySQLöÖÐþúÚ",
        'CLEAN(CHAR(9)&"Monthly report"&CHAR(10))': 'Monthly report',
    },

    CODE: {
        'CODE("C")': 67,
        'CODE("")': FormulaError.VALUE,
        // Unicode test
        'CODE("😀")': 128512,  // The first code point of emoji
        'CODE("中")': 20013,   // Chinese character
        'CODE("🚀")': 128640,  // Rocket emoji
        'CODE("👨‍👩‍👧‍👦")': 128104,  // Compound emoji
    },

    CONCAT: {
        'CONCAT(0, {1,2,3;5,6,7})': '0123567',
        'CONCAT(TRUE, 0, {1,2,3;5,6,7})': 'TRUE0123567',
        'CONCAT(0, {1,2,3;5,6,7},)': '0123567',
        'CONCAT("The"," ","sun"," ","will"," ","come"," ","up"," ","tomorrow.")': 'The sun will come up tomorrow.',
        'CONCAT({1,2,3}, "aaa", TRUE, 0, FALSE)': '123aaaTRUE0FALSE'
    },

    CONCATENATE: {
        'CONCATENATE({9,8,7})': '9',
        'CONCATENATE({9,8,7},{8,7,6})': '98',
        'CONCATENATE({9,8,7},"hello")': '9hello',
        'CONCATENATE({0,2,3}, 1, "A", TRUE, -12)': '01ATRUE-12',
    },

    DBCS: {
        'DBCS("ABC")': "ＡＢＣ",
        'DBCS("ｦｧｨｩ")': 'ヲァィゥ',
        'DBCS(",｡")': '，。',
    },

    DOLLAR: {
        'DOLLAR(1234567)': "$1,234,567.00",
        'DOLLAR(12345.67)': "$12,345.67"
    },

    EXACT: {
        'EXACT("hello", "hElLo")': false,
        'EXACT("HELLO","HELLO")': true
    },

    FIND: {
        'FIND("h","Hello")': FormulaError.VALUE,
        'FIND("o", "hello")': 5,
        // Unicode and emoji test
        'FIND("😁", "😀😁😂")': 2,
        'FIND("国", "中国加油")': 2,
        'FIND("世", "Hello世界")': 6,
        'FIND("🌟", "🚀🌟💫")': 2,
        'FIND("👨‍👩‍👧‍👦", "a👨‍👩‍👧‍👦b")': 2,  // Find compound emoji (starting from the 2nd character position in Excel)
    },

    FIXED: {
        'FIXED(1234.567, 1)': '1,234.6',
        'FIXED(12345.64123213)': '12,345.64',
        'FIXED(12345.64123213, 5)': '12,345.64123',
        'FIXED(12345.64123213, 5, TRUE)': '12345.64123',
        'FIXED(123456789.64, 5, FALSE)': '123,456,789.64000'
    },

    LEFT: {
        'LEFT("Salesman")': "S",
        'LEFT("Salesman",4)': "Sale",
        // Unicode and emoji test
        'LEFT("😀😁😂🤣", 2)': "😀😁",
        'LEFT("中国加油", 2)': "中国",
        'LEFT("Hello世界", 7)': "Hello世界",
        'LEFT("🚀🌟💫⭐", 1)': "🚀",
        'LEFT("👨‍👩‍👧‍👦👶", 1)': "👨‍👩‍👧‍👦",  // Compound emoji is treated as 1 character in Excel
    },

    RIGHT: {
        'RIGHT("Salseman",3)': "man",
        'RIGHT("Salseman")': "n",
        // Unicode and emoji test
        'RIGHT("😀😁😂🤣", 2)': "😂🤣",
        'RIGHT("中国加油", 2)': "加油",
        'RIGHT("Hello世界", 2)': "世界",
        'RIGHT("🚀🌟💫⭐", 1)': "⭐",
        'RIGHT("👶👨‍👩‍👧‍👦", 1)': "👨‍👩‍👧‍👦",  // Compound emoji is treated as 1 character in Excel
    },

    LEN: {
        'LEN("Phoenix, AZ")': 11,
        // Unicode and emoji test
        'LEN("😀😁😂")': 3,           // 3 separate emojis
        'LEN("中国")': 2,             // 2 Chinese characters
        'LEN("Hello世界")': 7,        // 5 English + 2 Chinese
        'LEN("🚀🌟💫")': 3,           // 3 emojis
        'LEN("👨‍👩‍👧‍👦")': 1,           // Compound emoji is treated as 1 character in Excel
        'LEN("a👨‍👩‍👧‍👦b")': 3,           // English + compound emoji (1) + English
    },

    LOWER: {
        'LOWER("E. E. Cummings")': "e. e. cummings",
        'LOWER("Apt. 2B")': "apt. 2b",
        'LOWER("HELLO WORLD")': "hello world",
        'LOWER("123Test")': "123test"
    },

    MID: {
        'MID("Fluid Flow",1,5)': "Fluid",
        'MID("Foo",5,1)': "",
        'MID("Foo",1,5)': "Foo",
        'MID("Foo",-1,5)': FormulaError.VALUE,
        'MID("Foo",1,-5)': FormulaError.VALUE,
        // Unicode and emoji test
        'MID("😀😁😂🤣", 2, 2)': "😁😂",
        'MID("中国加油", 2, 2)': "国加",
        'MID("Hello世界", 6, 2)': "世界",
        'MID("🚀🌟💫⭐", 2, 2)': "🌟💫",
        'MID("a👨‍👩‍👧‍👦b", 2, 1)': "👨‍👩‍👧‍👦",  // Extract compound emoji (treated as 1 character in Excel)
    },

    NUMBERVALUE: {
        'NUMBERVALUE("2.500,27",",",".")': 2500.27,
        // group separator occurs before the decimal separator
        'NUMBERVALUE("2500.,27",",",".")': 2500.27,
        'NUMBERVALUE("3.5%")': 0.035,
        'NUMBERVALUE("3 50")': 350,
        'NUMBERVALUE("$3 50")': 350,
        'NUMBERVALUE("($3 50)")': -350,
        'NUMBERVALUE("-($3 50)")': FormulaError.VALUE,
        'NUMBERVALUE("($-3 50)")': FormulaError.VALUE,
        'NUMBERVALUE("2500,.27",",",".")': FormulaError.VALUE,
        // group separator occurs after the decimal separator
        'NUMBERVALUE("3.5%",".",".")': FormulaError.VALUE,
        'NUMBERVALUE("3.5%",,)': FormulaError.VALUE,
        // decimal separator is used more than once
        'NUMBERVALUE("3..5")': FormulaError.VALUE,

    },

    PROPER: {
        'PROPER("this is a tiTle")': "This Is A Title",
        'PROPER("2-way street")': "2-Way Street",
        'PROPER("76BudGet")': '76Budget',
    },

    REPLACE: {
        'REPLACE("abcdefghijk",6,5,"*")': "abcde*k",
        'REPLACE("abcdefghijk",6,0,"*")': "abcde*fghijk",
        // Unicode and emoji tests
        'REPLACE("😀😁😂🤣", 2, 2, "🌟")': "😀🌟🤣",
        'REPLACE("中国加油", 2, 1, "华")': "中华加油",
        'REPLACE("Hello世界", 6, 2, "World")': "HelloWorld",
        'REPLACE("🚀🌟💫⭐", 2, 1, "🔥")': "🚀🔥💫⭐",
        'REPLACE("a👨‍👩‍👧‍👦b", 2, 1, "👶")': "a👶b",  // Replace compound emoji (treated as 1 character in Excel)
    },

    REPT: {
        'REPT("*_",4)': "*_*_*_*_"
    },

    SEARCH: {
        'SEARCH(",", "abcdef")': FormulaError.VALUE,
        'SEARCH("b", "abcdef")': 2,
        'SEARCH("c*f", "abcdef")': 3,
        'SEARCH("c?f", "abcdef")': FormulaError.VALUE,
        'SEARCH("c?e", "abcdef")': 3,
        'SEARCH("c\\b", "abcabcac\\bacb", 6)': 8,
        // Unicode and emoji tests
        'SEARCH("😁", "😀😁😂")': 2,
        'SEARCH("国", "中国加油")': 2,
        'SEARCH("世", "Hello世界")': 6,
        'SEARCH("🌟", "🚀🌟💫")': 2,
        'SEARCH("👨‍👩‍👧‍👦", "a👨‍👩‍👧‍👦b")': 2,  // Find compound emoji (starting from the 2nd character position in Excel)
    },

    SUBSTITUTE: {
        // Basic Function Test - Replace All Matches
        'SUBSTITUTE("销售数据", "销售", "成本")': "成本数据",
        'SUBSTITUTE("Hello World", "o", "0")': "Hell0 W0rld",
        
        // Specify which occurrence to replace
        'SUBSTITUTE("2008年第1季度", "1", "2", 1)': "2008年第2季度",
        'SUBSTITUTE("2011年第1季度", "1", "2", 2)': "2012年第1季度",
        'SUBSTITUTE("abcabc", "a", "X", 1)': "Xbcabc",
        'SUBSTITUTE("abcabc", "a", "X", 2)': "abcXbc",
        
        // Boundary case testing
        'SUBSTITUTE("hello", "", "world")': "hello",  // Empty old_text
        'SUBSTITUTE("hello", "x", "world")': "hello", // No match
        'SUBSTITUTE("hello", "hello", "")': "",       // Replace with empty string
        'SUBSTITUTE("", "a", "b")': "",               // Empty original text
        
        // Special character test
        'SUBSTITUTE("a.b.c", ".", "-")': "a-b-c",
        'SUBSTITUTE("a*b*c", "*", "+")': "a+b+c",
        'SUBSTITUTE("a[b]c", "[", "(")': "a(b]c",
        
        // Error testing
        'SUBSTITUTE("abc", "a", "X", 0)': FormulaError.VALUE,  // instance_num < 1
        'SUBSTITUTE("abc", "a", "X", -1)': FormulaError.VALUE, // instance_num < 1
        
        // Unicode and emoji tests
        'SUBSTITUTE("😀😁😀", "😀", "🌟")': "🌟😁🌟",        // Replace all emojis
        'SUBSTITUTE("😀😁😀", "😀", "🌟", 1)': "🌟😁😀",     // Replace first emoji
        'SUBSTITUTE("中国中国", "中", "华")': "华国华国",      // Replace Chinese characters
        'SUBSTITUTE("Hello世界World", "世界", "World")': "HelloWorldWorld", // Mixed Chinese and English
        'SUBSTITUTE("🚀🌟🚀💫", "🚀", "🔥")': "🔥🌟🔥💫",    // Replace emoji
    },

    T: {
        'T("*_")': "*_",
        'T(19)': "",
    },

    TEXT: {
        'TEXT(1234.567,"$#,##0.00")': "$1,234.57",
    },

    TEXTJOIN: {
        // Basic Function Test
        'TEXTJOIN(",", TRUE, "Apple", "Orange", "Banana")': "Apple,Orange,Banana",
        'TEXTJOIN("，", TRUE, "美元", "澳元", "人民币")': "美元，澳元，人民币",
        'TEXTJOIN("-", TRUE, "The", "sun", "will", "come", "up", "tomorrow")': "The-sun-will-come-up-tomorrow",
        
        // Empty delimiter test
        'TEXTJOIN("", TRUE, "A", "B", "C")': "ABC",
        
        // ignore_empty is FALSE, include empty values
        'TEXTJOIN(",", FALSE, "A", "", "B", "", "C")': "A,,B,,C",
        
        // ignore_empty is TRUE, ignore empty values
        'TEXTJOIN(",", TRUE, "A", "", "B", "", "C")': "A,B,C",
        
        // Array test
        'TEXTJOIN(",", TRUE, {"A", "B";"C", "D"})': "A,B,C,D",
        'TEXTJOIN("、", TRUE, {"a1", "b1"; "a2", "b2"})': "a1、b1、a2、b2",
        
        // Number delimiter test (numbers will be converted to text)
        'TEXTJOIN(1, TRUE, "A", "B", "C")': "A1B1C",
        
        // Single parameter test
        'TEXTJOIN(",", TRUE, "Single")': "Single",
        
        // Unicode and emoji tests
        'TEXTJOIN("🔥", TRUE, "Hello", "World")': "Hello🔥World",      // emoji as separator
        'TEXTJOIN(",", TRUE, "😀", "😁", "😂")': "😀,😁,😂",            // emoji content
        'TEXTJOIN("、", TRUE, "中国", "日本", "韩国")': "中国、日本、韩国", // Chinese content and separator
        'TEXTJOIN("", TRUE, "🚀", "🌟", "💫")': "🚀🌟💫",             // emoji without separator
    },

    TRIM: {
        'TRIM("     First Quarter Earnings    ")': "First Quarter Earnings"
    },

    UPPER: {
        'UPPER("e. e. cummings")': "E. E. CUMMINGS",
        'UPPER("total")': "TOTAL",
        'UPPER("Yield")': "YIELD",
        'UPPER("abc123")': "ABC123"
    },

    UNICHAR: {
        'UNICHAR(32)': " ",
        'UNICHAR(66)': "B",
        'UNICHAR(0)': FormulaError.VALUE,
        'UNICHAR(3333)': 'അ',
        // Unicode and emoji tests
        'UNICHAR(128512)': "😀",   // Happy emoji
        'UNICHAR(128640)': "🚀",   // Rocket emoji
        'UNICHAR(20013)': "中",    // Chinese character
        'UNICHAR(127775)': "🌟",   // Star emoji
        'UNICHAR(128104)': "👨",   // Man emoji (part of compound emoji)
    },

    UNICODE: {
        'UNICODE(" ")': 32,
        'UNICODE("B")': 66,
        'UNICODE("")': FormulaError.VALUE,
        // Unicode and emoji tests
        'UNICODE("😀")': 128512,   // Happy emoji
        'UNICODE("🚀")': 128640,   // Rocket emoji
        'UNICODE("中")': 20013,    // Chinese character
        'UNICODE("🌟")': 127775,   // Star emoji
        'UNICODE("👨‍👩‍👧‍👦")': 128104,  // First code point of compound emoji
    }

};
