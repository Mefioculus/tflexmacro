using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.IO;
using Xunit;
using FoxProShifrsNormalizer;


namespace FoxProShifrsNormalizerTests {
    public class FirstTest {

        public static List<object[]> Data() {
            Normalizer normalizer = new Normalizer();
            List<object[]> result = new List<object[]>();
            foreach (string shifr in File.ReadAllLines("shifrs.txt")) {
                result.Add(
                        new object[] { shifr,  normalizer.NormalizeShifrsFromFox(normalizer.NormalizeShifrsToFox(shifr))}
                        );
            }
            return result;
        }

        [Theory]
        [MemberData(nameof(Data))]
        public void TestNormalize(string expected, string processed) {
            Assert.Equal(expected, processed);
        }
    }
}
