/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class LanguageID
    {
        public const Int16 Afrikaans = 1078;
        public const Int16 Albanian = 1052;
        public const Int16 Amharic = 1118;
        public const Int16 ArabicAlgeria = 5121;
        public const Int16 ArabicBahrain = 15361;
        public const Int16 ArabicEgypt = 3073;
        public const Int16 ArabicIraq = 2049;
        public const Int16 ArabicJordan = 11265;
        public const Int16 ArabicKuwait = 13313;
        public const Int16 ArabicLebanon = 12289;
        public const Int16 ArabicLibya = 4097;
        public const Int16 ArabicMorocco = 6145;
        public const Int16 ArabicOman = 8193;
        public const Int16 ArabicQatar = 16385;
        public const Int16 ArabicSaudiArabia = 1025;
        public const Int16 ArabicSyria = 10241;
        public const Int16 ArabicTunisia = 7169;
        public const Int16 ArabicUAE = 14337;
        public const Int16 ArabicYemen = 9217;
        public const Int16 Armenian = 1067;
        public const Int16 Assamese = 1101;
        public const Int16 AzeriCyrillic = 2092;
        public const Int16 AzeriLatin = 1068;
        public const Int16 Basque = 1069;
        public const Int16 Belarusian = 1059;
        public const Int16 Bengali = 1093;
        public const Int16 BengaliBangladesh = 2117;
        public const Int16 Bulgarian = 1026;
        public const Int16 Burmese = 1109;
        public const Int16 Catalan = 1027;
        public const Int16 Cherokee = 1116;
        public const Int16 ChineseHongKong = 3076;
        public const Int16 ChineseMacao = 5124;
        public const Int16 ChinesePRC = 2052;
        public const Int16 ChineseSingapore = 4100;
        public const Int16 ChineseTaiwan = 1028;
        public const Int16 Croatian = 1050;
        public const Int16 Czech = 1029;
        public const Int16 Danish = 1030;
        public const Int16 Divehi = 1125;
        public const Int16 DutchBelgium = 2067;
        public const Int16 DutchNetherlands = 1043;
        public const Int16 Edo = 1126;
        public const Int16 EnglishAustralia = 3081;
        public const Int16 EnglishBelize = 10249;
        public const Int16 EnglishCanada = 4105;
        public const Int16 EnglishCaribbean = 9225;
        public const Int16 EnglishHongKong = 15369;
        public const Int16 EnglishIndia = 16393;
        public const Int16 EnglishIndonesia = 14345;
        public const Int16 EnglishIreland = 6153;
        public const Int16 EnglishJamaica = 8201;
        public const Int16 EnglishMalaysia = 17417;
        public const Int16 EnglishNewZealand = 5129;
        public const Int16 EnglishPhilippines = 13321;
        public const Int16 EnglishSingapore = 18441;
        public const Int16 EnglishSouthAfrica = 7177;
        public const Int16 EnglishTrinidadAndTobago = 11273;
        public const Int16 EnglishUK = 2057;
        public const Int16 EnglishUS = 1033;
        public const Int16 EnglishZimbabwe = 12297;
        public const Int16 Estonian = 1061;
        public const Int16 Faeroese = 1080;
        public const Int16 Farsi = 1065;
        public const Int16 Filipino = 1124;
        public const Int16 Finnish = 1035;
        public const Int16 FrenchBelgium = 2060;
        public const Int16 FrenchCameroon = 11276;
        public const Int16 FrenchCanada = 3084;
        public const Int16 FrenchCongoDRC = 9228;
        public const Int16 FrenchCotedIvoire = 12300;
        public const Int16 FrenchFrance = 1036;
        public const Int16 FrenchHaiti = 15372;
        public const Int16 FrenchLuxembourg = 5132;
        public const Int16 FrenchMali = 13324;
        public const Int16 FrenchMonaco = 6156;
        public const Int16 FrenchMorocco = 14348;
        public const Int16 FrenchReunion = 8204;
        public const Int16 FrenchSenegal = 10252;
        public const Int16 FrenchSwitzerland = 4108;
        public const Int16 FrenchWestIndies = 7180;
        public const Int16 FrisianNetherlands = 1122;
        public const Int16 Fulfulde = 1127;
        public const Int16 FYROMacedonian = 1071;
        public const Int16 GaelicIreland = 2108;
        public const Int16 GaelicScotland = 1084;
        public const Int16 Galician = 1110;
        public const Int16 Georgian = 1079;
        public const Int16 GermanAustria = 3079;
        public const Int16 GermanGermany = 1031;
        public const Int16 GermanLiechtenstein = 5127;
        public const Int16 GermanLuxembourg = 4103;
        public const Int16 GermanSwitzerland = 2055;
        public const Int16 Greek = 1032;
        public const Int16 Guarani = 1140;
        public const Int16 Gujarati = 1095;
        public const Int16 Hausa = 1128;
        public const Int16 Hawaiian = 1141;
        public const Int16 Hebrew = 1037;
        public const Int16 Hindi = 1081;
        public const Int16 Hungarian = 1038;
        public const Int16 Ibibio = 1129;
        public const Int16 Icelandic = 1039;
        public const Int16 Igbo = 1136;
        public const Int16 Indonesian = 1057;
        public const Int16 Inuktitut = 1117;
        public const Int16 ItalianItaly = 1040;
        public const Int16 ItalianSwitzerland = 2064;
        public const Int16 Japanese = 1041;
        public const Int16 Kannada = 1099;
        public const Int16 Kanuri = 1137;
        public const Int16 Kashmiri = 2144;
        public const Int16 KashmiriArabic = 1120;
        public const Int16 Kazakh = 1087;
        public const Int16 Khmer = 1107;
        public const Int16 Konkani = 1111;
        public const Int16 Korean = 1042;
        public const Int16 Kyrgyz = 1088;
        public const Int16 Lao = 1108;
        public const Int16 Latin = 1142;
        public const Int16 Latvian = 1062;
        public const Int16 Lithuanian = 1063;
        public const Int16 Malay = 1086;
        public const Int16 MalayBruneiDarussalam = 2110;
        public const Int16 Malayalam = 1100;
        public const Int16 Maltese = 1082;
        public const Int16 Manipuri = 1112;
        public const Int16 Maori = 1153;
        public const Int16 Marathi = 1102;
        public const Int16 Mongolian = 1104;
        public const Int16 MongolianMongolian = 2128;
        public const Int16 Nepali = 1121;
        public const Int16 NepaliIndia = 2145;
        public const Int16 NorwegianBokm�l = 1044;
        public const Int16 NorwegianNynorsk = 2068;
        public const Int16 Oriya = 1096;
        public const Int16 Oromo = 1138;
        public const Int16 Papiamentu = 1145;
        public const Int16 Pashto = 1123;
        public const Int16 Polish = 1045;
        public const Int16 PortugueseBrazil = 1046;
        public const Int16 PortuguesePortugal = 2070;
        public const Int16 Punjabi = 1094;
        public const Int16 PunjabiPakistan = 2118;
        public const Int16 QuechuaBolivia = 1131;
        public const Int16 QuechuaEcuador = 2155;
        public const Int16 QuechuaPeru = 3179;
        public const Int16 RhaetoRomanic = 1047;
        public const Int16 RomanianMoldova = 2072;
        public const Int16 RomanianRomania = 1048;
        public const Int16 RussianMoldova = 2073;
        public const Int16 RussianRussia = 1049;
        public const Int16 SamiLappish = 1083;
        public const Int16 Sanskrit = 1103;
        public const Int16 Sepedi = 1132;
        public const Int16 SerbianCyrillic = 3098;
        public const Int16 SerbianLatin = 2074;
        public const Int16 SindhiArabic = 2137;
        public const Int16 SindhiDevanagari = 1113;
        public const Int16 Sinhalese = 1115;
        public const Int16 Slovak = 1051;
        public const Int16 Slovenian = 1060;
        public const Int16 Somali = 1143;
        public const Int16 Sorbian = 1070;
        public const Int16 SpanishArgentina = 11274;
        public const Int16 SpanishBolivia = 16394;
        public const Int16 SpanishChile = 13322;
        public const Int16 SpanishColombia = 9226;
        public const Int16 SpanishCostaRica = 5130;
        public const Int16 SpanishDominicanRepublic = 7178;
        public const Int16 SpanishEcuador = 12298;
        public const Int16 SpanishElSalvador = 17418;
        public const Int16 SpanishGuatemala = 4106;
        public const Int16 SpanishHonduras = 18442;
        public const Int16 SpanishMexico = 2058;
        public const Int16 SpanishNicaragua = 19466;
        public const Int16 SpanishPanama = 6154;
        public const Int16 SpanishParaguay = 15370;
        public const Int16 SpanishPeru = 10250;
        public const Int16 SpanishPuertoRico = 20490;
        public const Int16 SpanishSpainModernSort = 3082;
        public const Int16 SpanishSpainTraditionalSort = 1034;
        public const Int16 SpanishUruguay = 14346;
        public const Int16 SpanishVenezuela = 8202;
        public const Int16 Sutu = 1072;
        public const Int16 Swahili = 1089;
        public const Int16 SwedishFinland = 2077;
        public const Int16 SwedishSweden = 1053;
        public const Int16 Syriac = 1114;
        public const Int16 Tajik = 1064;
        public const Int16 Tamazight = 1119;
        public const Int16 TamazightLatin = 2143;
        public const Int16 Tamil = 1097;
        public const Int16 Tatar = 1092;
        public const Int16 Telugu = 1098;
        public const Int16 Thai = 1054;
        public const Int16 TibetanBhutan = 2129;
        public const Int16 TibetanPRC = 1105;
        public const Int16 TigrignaEritrea = 2163;
        public const Int16 TigrignaEthiopia = 1139;
        public const Int16 Tsonga = 1073;
        public const Int16 Tswana = 1074;
        public const Int16 Turkish = 1055;
        public const Int16 Turkmen = 1090;
        public const Int16 Ukrainian = 1058;
        public const Int16 Urdu = 1056;
        public const Int16 UzbekCyrillic = 2115;
        public const Int16 UzbekLatin = 1091;
        public const Int16 Venda  = 1075;
        public const Int16 Vietnamese = 1066;
        public const Int16 Welsh = 1106;
        public const Int16 Xhosa = 1076;
        public const Int16 Yi = 1144;
        public const Int16 Yiddish = 1085;
        public const Int16 Yoruba = 1130;
        public const Int16 Zulu = 1077;
    }
}
