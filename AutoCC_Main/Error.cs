using System.Collections.Generic;

namespace AutoCC_Main
{
    class Error
    {

        // 일러스트 오류 나면 메인에서 일러스트 제시작
        public const string Com_Exception = "XXXX";


        //Time Out 2020.05.04
        public const string Time_Out = "9999";


        //예외업체 추가 2019.04.23
        public const string exclude_customer = "0000";

        // system
        public const string FILE_NOT_FOUND = "0101";
        public const string FILE_EXT_NOT_EXIST = "0102";

        // general
        public const string FAILED_TO_EXECUTE_DRAWER = "0201";
        public const string EXCEPTION_OCCURRED_IN_DRAWER = "0202";
        public const string FAILED_TO_CREATE_DOCUMENT = "0211";
        public const string FAILED_TO_OPEN_DOCUMENT = "0212";
        public const string DRAWER_VERSION_NOT_ALLOWED = "0213";
        public const string FAILED_TO_CREATE_PDF = "0214";
        public const string FAILED_TO_COPY_PDF = "0215";
        public const string BLANK_PDF = "0216";
        public const string TIMEOUT_TO_INSPECT = "0221";
        public const string INVALID_WORK_DATA = "0222";
        public const string FILE_EXT_NOT_ALLOWED = "0231";
        public const string CATEGORY_NOT_ALLOWED = "0232";
        public const string DOCUMENT_MADE_IN_MAC_OS = "0233";
        public const string FAILED_TO_DISPATCH_DATABASE = "0241";

        // page
        public const string PAGE_ERROR_UNKNOWN = "0300";
        public const string PAGE_NOT_EXIST = "0301";
        public const string PAGE_COUNT_NOT_MATCH_TO_DOUBLE_SIDE = "0302";
        public const string PAGE_COUNT_NOT_MATCH_TO_ITEM_COUNT = "0303";
        public const string OVERLAPPED_PAGES_EXIST = "0304";
        public const string FAILED_TO_PAIR_PAGES = "0305";
        public const string BLEED_RECT_NOT_EXIST = "0306";
        public const string REVERSED_PAGE_ASPECT = "0307";
        public const string FAILED_TO_ERASE_OUTLINE = "0311";
        public const string DUMMY_OBJECT_IN_PAGE = "0312";

        // [validation] document
        public const string DOCUMENT_ERROR_UNKNOWN = "0400";
        public const string DOCUMENT_WITH_NOT_A_LAYER = "0401";
        public const string DOCUMENT_WITH_NOT_A_PAGE = "0402"; // 코렐 페이지가 1개이상 일떄
        public const string DOCUMENT_WITH_RGB = "0403";
        public const string DOCUMENT_WITH_TOO_MANY_OBJECT = "0404";

        // [validation] object
        public const string TEXT_OBJECT_USED = "0501";
        public const string OBJECT_WITH_RGB = "0502";
        public const string OBJECT_WITH_LAB = "0503";
        public const string OBJECT_WITH_LINK = "0504";
        public const string OBJECT_WITH_PDF_LINK = "0505";
        public const string OBJECT_WITH_NOT_CMYK = "0507";
        public const string OBJECT_WITH_SPOT_COLOR = "0508";
        public const string OBJECT_WITH_FILL_FOUNTAIN = "0511"; // 2019-01-21 여기서 error
        public const string OBJECT_WITH_FILL_PATTERN = "0512";
        public const string OBJECT_WITH_FILL_TEXTURE = "0513";
        public const string OBJECT_WITH_FILL_POSTSCRIPT = "0514";
        public const string OBJECT_WITH_FILL_HATCH = "0515";
        public const string OBJECT_WITH_FILL_LOWER_TONE = "0516";
        public const string OBJECT_WITH_STROKE_TOO_THICK = "0521";
        public const string OBJECT_WITH_LOWER_DPI = "0531";
        public const string OBJECT_WITH_UPPER_DPI = "0532";

        public const string OBJECT_Raster_Image = "0533"; //2020.04.29 추가
        public const string OBJECT_Count_Small = "0534"; //2020.05.06 추가 객체수가 모자람


        public const string OBJECT_Opacity_Used = "0535"; //2020.05.15 추가
        public const string OBJECT_PatternColor_Used = "0536";  //2020.05.15 추가
        public const string OBJECT_RGBColor_Used = "0537";  //2020.05.15 추가
        public const string OBJECT_LabColor_Used = "0538";  //2020.05.15 추가
        public const string OBJECT_Texture_Used = "0539";  //2020.05.15 추가





        // [validation] effect
        public const string EFFECT_CONTOUR_USED = "0601";
        public const string EFFECT_TRANSPARENCY_USED = "0602";
        public const string EFFECT_BLEND_USED = "0603";
        public const string EFFECT_CUSTOM_USED = "0604";
        public const string EFFECT_DISTORTION_USED = "0605";
        public const string EFFECT_DROPSHADOW_USED = "0606";
        public const string EFFECT_ENVELOPE_USED = "0607";
        public const string EFFECT_EXTRUDE_USED = "0608";
        public const string EFFECT_PERSPECTIVE_USED = "0609";
        public const string EFFECT_BEVEL_USED = "0610";
        public const string EFFECT_POWERCLIP_USED = "0611";
        public const string EFFECT_LENS_USED = "0612";
        public const string EFFECT_CONTROLPATH_USED = "0613";

        public static string ToDesc(string error_code, string lang_id = "")     // en, ko. default: en. 
        {
            if (lang_id == "")
                lang_id = "en";

            string desc = string.Empty;

            switch (error_code)
            {
                case FILE_NOT_FOUND:
                    if (lang_id == "en")
                        desc = "FILE_NOT_FOUND";
                    else if (lang_id == "ko")
                        desc = "file이 존재하지 않는다.";
                    break;
                case FILE_EXT_NOT_EXIST:
                    if (lang_id == "en")
                        desc = "FILE_EXT_NOT_EXIST";
                    else if (lang_id == "ko")
                        desc = "file의 ext가 없다.";
                    break;
                case FILE_EXT_NOT_ALLOWED:
                    if (lang_id == "en")
                        desc = "FILE_EXT_NOT_ALLOWED";
                    else if (lang_id == "ko")
                        desc = "허용된 ext가 아니다.";
                    break;
                case CATEGORY_NOT_ALLOWED:
                    if (lang_id == "en")
                        desc = "CATEGORY_NOT_ALLOWED";
                    else if (lang_id == "ko")
                        desc = "허용된 상품종류가 아니다.";
                    break;
                case DOCUMENT_MADE_IN_MAC_OS:
                    if (lang_id == "en")
                        desc = "DOCUMENT_MADE_IN_MAC_OS";
                    else if (lang_id == "ko")
                        desc = "IBM OS용 문서가 아니다.";
                    break;
                case FAILED_TO_DISPATCH_DATABASE:
                    if (lang_id == "en")
                        desc = "FAILED_TO_DISPATCH_DATABASE";
                    else if (lang_id == "ko")
                        desc = "DATABASE 처리에 실패";
                    break;
                case FAILED_TO_EXECUTE_DRAWER:
                    if (lang_id == "en")
                        desc = "FAILED_TO_EXECUTE_DRAWER";
                    else if (lang_id == "ko")
                        desc = "drawer 실행 실패";
                    break;
                case EXCEPTION_OCCURRED_IN_DRAWER:
                    if (lang_id == "en")
                        desc = "EXCEPTION_OCCURRED_IN_DRAWER";
                    else if (lang_id == "ko")
                        desc = "drawer에서 exception 발생";
                    break;
                case FAILED_TO_CREATE_DOCUMENT:
                    if (lang_id == "en")
                        desc = "FAILED_TO_CREATE_DOCUMENT";
                    else if (lang_id == "ko")
                        desc = "문서 생성 실패";
                    break;
                case FAILED_TO_OPEN_DOCUMENT:
                    if (lang_id == "en")
                        desc = "FAILED_TO_OPEN_DOCUMENT";
                    else if (lang_id == "ko")
                        desc = "문서 열기 실패";
                    break;
                case DRAWER_VERSION_NOT_ALLOWED:
                    if (lang_id == "en")
                        desc = "DRAWER_VERSION_NOT_ALLOWED";
                    else if (lang_id == "ko")
                        desc = "지원하지 않는 버전의 문서";
                    break;
                case FAILED_TO_CREATE_PDF:
                    if (lang_id == "en")
                        desc = "FAILED_TO_CREATE_PDF";
                    else if (lang_id == "ko")
                        desc = "PDF 생성 실패";
                    break;
                case FAILED_TO_COPY_PDF:
                    if (lang_id == "en")
                        desc = "FAILED_TO_COPY_PDF";
                    else if (lang_id == "ko")
                        desc = "PDF 전송 실패";
                    break;
                case BLANK_PDF:
                    if (lang_id == "en")
                        desc = "BLANK_PDF";
                    else if (lang_id == "ko")
                        desc = "빈 PDF";
                    break;
                case TIMEOUT_TO_INSPECT:
                    if (lang_id == "en")
                        desc = "TIMEOUT_TO_INSPECT";
                    else if (lang_id == "ko")
                        desc = "점검 시간 초과";
                    break;
                case INVALID_WORK_DATA:
                    if (lang_id == "en")
                        desc = "INVALID_WORK_DATA";
                    else if (lang_id == "ko")
                        desc = "유효하지 않은 주문 데이터";
                    break;
                case PAGE_NOT_EXIST:
                    if (lang_id == "en")
                        desc = "PAGE_NOT_EXIST";
                    else if (lang_id == "ko")
                        desc = "page가 존재하지 않는다.";
                    break;
                case PAGE_COUNT_NOT_MATCH_TO_DOUBLE_SIDE:
                    if (lang_id == "en")
                        desc = "PAGE_COUNT_NOT_MATCH_TO_DOUBLE_SIDE";
                    else if (lang_id == "ko")
                        desc = "page 수가 양면에 부합하지 않는다. (홀수이다.)";
                    break;
                case PAGE_COUNT_NOT_MATCH_TO_ITEM_COUNT:
                    if (lang_id == "en")
                        desc = "PAGE_COUNT_NOT_MATCH_TO_ITEM_COUNT";
                    else if (lang_id == "ko")
                        desc = "page 수가 건수에 부합하지 않는다.";
                    break;
                case OVERLAPPED_PAGES_EXIST:
                    if (lang_id == "en")
                        desc = "OVERLAPPED_PAGES_EXIST";
                    else if (lang_id == "ko")
                        desc = "page들간에 겹침이 있다.";
                    break;
                case FAILED_TO_PAIR_PAGES:
                    if (lang_id == "en")
                        desc = "FAILED_TO_PAIR_PAGES";
                    else if (lang_id == "ko")
                        desc = "page들의 배치가 일정하지 않아서, 양면 묶음을 만들수가 없다.";
                    break;
                case BLEED_RECT_NOT_EXIST:
                    if (lang_id == "en")
                        desc = "BLEED_RECT_NOT_EXIST";
                    else if (lang_id == "ko")
                        desc = "작업사이즈는 없고, 재단사이즈만 있다.";
                    break;
                case REVERSED_PAGE_ASPECT:
                    if (lang_id == "en")
                        desc = "REVERSED_PAGE_ASPECT";
                    else if (lang_id == "ko")
                        desc = "작업사이즈의 가로/세로값이 바뀌어서 있는 page들만 존재한다.";
                    break;
                case FAILED_TO_ERASE_OUTLINE:
                    if (lang_id == "en")
                        desc = "FAILED_TO_ERASE_OUTLINE";
                    else if (lang_id == "ko")
                        desc = "외곽선 삭제 실패";
                    break;
                case DUMMY_OBJECT_IN_PAGE:
                    if (lang_id == "en")
                        desc = "DUMMY_OBJECT_IN_PAGE";
                    else if (lang_id == "ko")
                        desc = "페이지 내에 더비 개체가 있다.";
                    break;
                case DOCUMENT_WITH_NOT_A_LAYER:
                    if (lang_id == "en")
                        desc = "DOCUMENT_WITH_NOT_A_LAYER";
                    else if (lang_id == "ko")
                        desc = "layer count가 1이 아니다.";
                    break;
                case DOCUMENT_WITH_NOT_A_PAGE:
                    if (lang_id == "en")
                        desc = "DOCUMENT_WITH_NOT_A_PAGE";
                    else if (lang_id == "ko")
                        desc = "페이지가 한장 이상 있다.";
                    break;
                case DOCUMENT_WITH_RGB:
                    if (lang_id == "en")
                        desc = "DOCUMENT_WITH_RGB";
                    else if (lang_id == "ko")
                        desc = "문서에 RGB 사용";
                    break;
                case DOCUMENT_WITH_TOO_MANY_OBJECT:
                    if (lang_id == "en")
                        desc = "DOCUMENT_WITH_TOO_MANY_OBJECT";
                    else if (lang_id == "ko")
                        desc = "객체 n개 이상";
                    break;
                case TEXT_OBJECT_USED:
                    if (lang_id == "en")
                        desc = "TEXT_OBJECT_USED";
                    else if (lang_id == "ko")
                        desc = "텍스트개체 사용";
                    break;
                case OBJECT_WITH_RGB:
                    if (lang_id == "en")
                        desc = "OBJECT_WITH_RGB";
                    else if (lang_id == "ko")
                        desc = "개체에 rgb 칼라 사용";
                    break;
                case OBJECT_WITH_LAB:
                    if (lang_id == "en")
                        desc = "OBJECT_WITH_LAB";
                    else if (lang_id == "ko")
                        desc = "개체에 lab 칼라 사용";
                    break;
                case OBJECT_WITH_LINK:
                    if (lang_id == "en")
                        desc = "OBJECT_WITH_LINK";
                    else if (lang_id == "ko")
                        desc = "링크이미지 사용";
                    break;
                case OBJECT_WITH_PDF_LINK:
                    if (lang_id == "en")
                        desc = "OBJECT_WITH_PDF_LINK";
                    else if (lang_id == "ko")
                        desc = "PDF링크이미지 사용";
                    break;
                case OBJECT_WITH_LOWER_DPI:
                    if (lang_id == "en")
                        desc = "OBJECT_WITH_LOWER_DPI";
                    else if (lang_id == "ko")
                        desc = "사용된 image 개체의 DPI가 낮다.";
                    break;
                case OBJECT_WITH_UPPER_DPI:
                    if (lang_id == "en")
                        desc = "OBJECT_WITH_UPPER_DPI";
                    else if (lang_id == "ko")
                        desc = "사용된 image 개체의 DPI가 높다.";
                    break;
                case OBJECT_WITH_NOT_CMYK:
                    if (lang_id == "en")
                        desc = "OBJECT_WITH_NOT_CMYK";
                    else if (lang_id == "ko")
                        desc = "개체에 cmyk 칼라를 사용하지 않음";
                    break;
                case OBJECT_WITH_SPOT_COLOR:
                    if (lang_id == "en")
                        desc = "OBJECT_WITH_SPOT_COLOR";
                    else if (lang_id == "ko")
                        desc = "개체에 spot 칼라 사용";
                    break;
                case OBJECT_WITH_FILL_FOUNTAIN:
                    if (lang_id == "en")
                        desc = "OBJECT_WITH_FILL_FOUNTAIN";
                    else if (lang_id == "ko")
                        desc = "채움: 계조채움";
                    break;
                case OBJECT_WITH_FILL_PATTERN:
                    if (lang_id == "en")
                        desc = "OBJECT_WITH_FILL_PATTERN";
                    else if (lang_id == "ko")
                        desc = "채움: 무늬채움";
                    break;
                case OBJECT_WITH_FILL_TEXTURE:
                    if (lang_id == "en")
                        desc = "OBJECT_WITH_FILL_TEXTURE";
                    else if (lang_id == "ko")
                        desc = "채움: 텍스쳐채움";
                    break;
                case OBJECT_WITH_FILL_POSTSCRIPT:
                    if (lang_id == "en")
                        desc = "OBJECT_WITH_FILL_POSTSCRIPT";
                    else if (lang_id == "ko")
                        desc = "채움: PostScript채움";
                    break;
                case OBJECT_WITH_FILL_HATCH:
                    if (lang_id == "en")
                        desc = "OBJECT_WITH_FILL_HATCH";
                    else if (lang_id == "ko")
                        desc = "채움: 해치채움";
                    break;
                case OBJECT_WITH_FILL_LOWER_TONE:
                    if (lang_id == "en")
                        desc = "OBJECT_WITH_FILL_LOWER_TONE";
                    else if (lang_id == "ko")
                        desc = "채움: 낮은 색조";
                    break;
                case OBJECT_WITH_STROKE_TOO_THICK:
                    if (lang_id == "en")
                        desc = "OBJECT_WITH_STROKE_TOO_THICK";
                    else if (lang_id == "ko")
                        desc = "윤곽선: 두꺼움";
                    break;
                case EFFECT_CONTOUR_USED:
                    if (lang_id == "en")
                        desc = "EFFECT_CONTOUR_USED";
                    else if (lang_id == "ko")
                        desc = "윤곽효과";
                    break;
                case EFFECT_TRANSPARENCY_USED:
                    if (lang_id == "en")
                        desc = "EFFECT_TRANSPARENCY_USED";
                    else if (lang_id == "ko")
                        desc = "투명효과";
                    break;
                case EFFECT_BLEND_USED:
                    if (lang_id == "en")
                        desc = "EFFECT_BLEND_USED";
                    else if (lang_id == "ko")
                        desc = "블랜드효과";
                    break;
                case EFFECT_CUSTOM_USED:
                    if (lang_id == "en")
                        desc = "EFFECT_CUSTOM_USED";
                    else if (lang_id == "ko")
                        desc = "custom효과";
                    break;
                case EFFECT_DISTORTION_USED:
                    if (lang_id == "en")
                        desc = "EFFECT_DISTORTION_USED";
                    else if (lang_id == "ko")
                        desc = "비틀기효과";
                    break;
                case EFFECT_DROPSHADOW_USED:
                    if (lang_id == "en")
                        desc = "EFFECT_DROPSHADOW_USED";
                    else if (lang_id == "ko")
                        desc = "그림자효과";
                    break;
                case EFFECT_ENVELOPE_USED:
                    if (lang_id == "en")
                        desc = "EFFECT_ENVELOPE_USED";
                    else if (lang_id == "ko")
                        desc = "형태효과";
                    break;
                case EFFECT_EXTRUDE_USED:
                    if (lang_id == "en")
                        desc = "EFFECT_EXTRUDE_USED";
                    else if (lang_id == "ko")
                        desc = "입체효과";
                    break;
                case EFFECT_PERSPECTIVE_USED:
                    if (lang_id == "en")
                        desc = "EFFECT_PERSPECTIVE_USED";
                    else if (lang_id == "ko")
                        desc = "원근감";
                    break;
                case EFFECT_BEVEL_USED:
                    if (lang_id == "en")
                        desc = "EFFECT_BEVEL_USED";
                    else if (lang_id == "ko")
                        desc = "베벨효과";
                    break;
                case EFFECT_POWERCLIP_USED:
                    if (lang_id == "en")
                        desc = "EFFECT_POWERCLIP_USED";
                    else if (lang_id == "ko")
                        desc = "파워클립";
                    break;
                case EFFECT_LENS_USED:
                    if (lang_id == "en")
                        desc = "EFFECT_LENS_USED";
                    else if (lang_id == "ko")
                        desc = "렌즈효과";
                    break;
                case EFFECT_CONTROLPATH_USED:
                    if (lang_id == "en")
                        desc = "EFFECT_CONTROLPATH_USED";
                    else if (lang_id == "ko")
                        desc = "도안매체효과";
                    break;
                default:
                    desc = "";
                    break;
            }

            return desc;
        }

        public static void AddErrorCode(string error_code, bool repair, ref List<string> error_codes)
        {
            if (error_codes == null)
                error_codes = new List<string>();

            if (error_codes.Contains(error_code) == false)
            {
                string log = string.Format("[AddErrorCode] {0} (repair: {1})", Error.ToDesc(error_code), repair.ToString());

                error_codes.Add(error_code);
            }
        }

        public static void DeleteErrorCode(string error_code, ref List<string> error_codes)
        {
            if (error_codes == null)
                return;

            if (error_codes.Contains(error_code) == true)
            {
                string log = string.Format("[DeleteErrorCode] {0}", Error.ToDesc(error_code));

                error_codes.Remove(error_code);
            }
        }
    }
}