// The MIT License(MIT)
// Copyright(c) 2022 mocchi
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files(the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions :
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

#include "zipper.h"
#include "json.hpp"
#include <cstdint>

namespace cpp_pptx_template {
	struct image_data {
		uint8_t *filedata; /// 画像データ本体  呼び出し元でメモリ管理すること
		size_t filedata_size; /// 画像データサイズ
		double aspect_y_x; /// 縦横比(縦幅 ÷ 横幅)
		const char *ext; /// 画像データの拡張子
	};

	/// template_pptx : 雛形PowerPoint(zipper::UnZipに入れて呼び出す)
	/// pptx_to_generate: PowerPoint生成先(zipper::Zipに書き出す)
	/// slides_define: スライド定義json
	/// image_data_ary: 画像データ配列
	/// image_data_count: 画像要素数
	void create_slide_from_template(zipper::UnZip &template_pptx, zipper::Zip &pptx_to_generate, const nlohmann::json &slides_define, const image_data *image_data_ary, int image_data_count);
}
