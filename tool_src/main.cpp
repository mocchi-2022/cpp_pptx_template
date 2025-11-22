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

#include <cstdio>
#include <string>
#include <vector>
#include <list>
#include "zipper.h"
#include "json.hpp"
#include "cpp_pptx_template.h"

#define STB_IMAGE_IMPLEMENTATION
#define STBI_NO_PSD
#define STBI_NO_TGA
#define STBI_NO_HDR
#define STBI_NO_PIC
#define STBI_NO_PNM
#include "stb_image.h"

template<typename C> void read_file(const char *path, C &filedata) {
	FILE *fp = std::fopen(path, "rb");
	std::fseek(fp, 0, SEEK_END);
	size_t sz = std::ftell(fp);
	std::rewind(fp);
	filedata.resize(sz);
	std::fread(&filedata[0], sz, 1, fp);
	std::fclose(fp);
}

void chomp(char *line) {
	char *pl = line + std::strlen(line) - 1;
	while (pl > line && *pl == '\r' || *pl == '\n') *pl = '\0', --pl;
}

struct zlib_memfunc64_def : public zlib_filefunc64_def_s {
	std::vector<uint8_t> mem;
	ZPOS64_T fpos;
	zlib_memfunc64_def() : fpos(0L){
		zopen64_file = [](voidpf opaque, const void* filename, int mode) {
			return opaque;
		};
		zread_file = [](voidpf opaque, voidpf stream, void* buf, uLong size) -> uLong {
			zlib_memfunc64_def *this_ = static_cast<zlib_memfunc64_def *>(opaque);
			if (!this_) return 0;
			uLong size_toread = size;
			if (this_->mem.size() <= this_->fpos) return 0;
			ZPOS64_T remain = this_->mem.size() - this_->fpos;
			if (static_cast<ZPOS64_T>(size_toread) > remain) size_toread = static_cast<uLong>(remain);
			std::memcpy(buf, this_->mem.data() + this_->fpos, size_toread);
			this_->fpos += size_toread;
			return size_toread;
		};
		zwrite_file = [](voidpf opaque, voidpf stream, const void* buf, uLong size) -> uLong {
			zlib_memfunc64_def *this_ = static_cast<zlib_memfunc64_def *>(opaque);
			if (!this_) return 0;
			if (this_->mem.size() != this_->fpos) this_->mem.resize(this_->fpos);
			const uint8_t *buf_u8 = static_cast<const uint8_t *>(buf);
			this_->mem.insert(this_->mem.end(), buf_u8, buf_u8 + size);
			this_->fpos += size;
			return size;
		};
		ztell64_file = [](voidpf opaque, voidpf stream) -> ZPOS64_T{
			zlib_memfunc64_def *this_ = static_cast<zlib_memfunc64_def *>(opaque);
			if (!this_) return 0;
			return this_->fpos;
		};
		zseek64_file = [](voidpf opaque, voidpf stream, ZPOS64_T offset, int origin) -> long {
			zlib_memfunc64_def *this_ = static_cast<zlib_memfunc64_def *>(opaque);
			if (!this_) return 1;
			if (origin == ZLIB_FILEFUNC_SEEK_SET) {
				this_->fpos = offset;
			} else if (origin == ZLIB_FILEFUNC_SEEK_CUR) {
				this_->fpos += offset;
			} else if (origin == ZLIB_FILEFUNC_SEEK_END) {
				this_->fpos = this_->mem.size() + offset;
			} else return 1;
			return 0;
		};
		zclose_file = [](voidpf opaque, voidpf stream)-> int {
			return 0;
		};
		zerror_file = [](voidpf opaque, voidpf stream) -> int {
			return 0;
		};
		opaque = this;
	}
};

int main(int argc, char *argv[]) {
	if (argc != 5) {
		std::printf("usage: cpp_pptx_template.exe tempalte_pptx pptx_to_generate slides_define_json imagelist");
		return 0;
	}

	nlohmann::json slides_define;

	{
		std::string json_str;
		read_file(argv[3], json_str);
		slides_define = nlohmann::json::parse(json_str);
	}

	std::vector<cpp_pptx_template::image_data> image_data_ary;
	std::list<std::vector<uint8_t>> image_filedata;
	{
		FILE *fp = std::fopen(argv[4], "r");
		char line[256];
		const char *png_ext = "png";
		for (;;) {
			line[0] = '\0';
			std::fgets(line, 255, fp);
			chomp(line);
			if (line[0] != '\0') {
				image_filedata.push_back(std::vector<uint8_t>());
				auto &imgfile = image_filedata.back();
				read_file(line, imgfile);
				image_data_ary.push_back(cpp_pptx_template::image_data());
				auto &image_data = image_data_ary.back();
				image_data.filedata = imgfile.data();
				image_data.filedata_size = imgfile.size();
				image_data.ext = png_ext;
				int width, height, channels;
				auto stbi = stbi_load_from_memory(image_data.filedata, static_cast<int>(image_data.filedata_size), &width, &height, &channels, 0);
				stbi_image_free(stbi);
				image_data.aspect_y_x = static_cast<double>(height) / static_cast<double>(width);
			}
			if (std::feof(fp)) break;
		}
		std::fclose(fp);
	}

	{
		zlib_memfunc64_def memfile;
		read_file(argv[1], memfile.mem);
		zipper::UnZip template_pptx("", &memfile);
		zipper::Zip pptx_to_generate(argv[2]);

		create_slide_from_template(template_pptx, pptx_to_generate, slides_define, image_data_ary.data(), static_cast<int>(image_data_ary.size()));
	}

	return 0;
}