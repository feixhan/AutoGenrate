using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using CppRipper;

namespace AutoGenrate
{
    class TestFile
    {
        public static void genNewFile(string curFileName, List<FunctionDefine> functionList)
        {
            string newFileName = Path.GetFileNameWithoutExtension(curFileName);
            //1.header file
            string fileContent = "#include <embUnit/embUnit.h>\r\n\r\n";
            //2.make function
            fileContent += ("static void setUp(void)\r\n"
                        + "{\r\n"
                        + "\t/* initialize */\r\n"
                        + "}\r\n\r\n"
                        + "static void tearDown(void)\r\n"
                        + "{\r\n"
                        + "\t/* terminate */\r\n"
                        + "}\r\n\r\n");
            foreach (FunctionDefine item in functionList)
            {
                fileContent += ("static void " + item.FunctionName + "(void)\r\n {\r\n"
                            + "	/* please delete below code and  add you code here!! */\r\n"
                            + "	TEST_FAIL(\"no implementation\");\r\n"
                            + "}\r\n\r\n");
            }

            fileContent += ("TestRef " + newFileName + "_tests(void)\r\n"
                        + "{\r\n"
                        + "	EMB_UNIT_TESTFIXTURES(fixtures) {\r\n");

            foreach (FunctionDefine item in functionList)
            {
                fileContent += string.Format("		new_TestFixture(\x22{0}\x22,{1}),\r\n", item.FunctionName, item.FunctionName);
            }
            fileContent += ("	};\r\n"
                        + string.Format("	EMB_UNIT_TESTCALLER({0},\x22{0}\x22,setUp,tearDown,fixtures);\r\n", newFileName)
                        + string.Format("	return (TestRef)&{0};\r\n", newFileName)
                        + "};\r\n");

            FileStream fs = new FileStream(curFileName, FileMode.CreateNew);
            StreamWriter sw = new StreamWriter(fs);
            sw.Write(fileContent);
            sw.Flush();
            sw.Close();
            fs.Close();
        }
    }
}
