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
#include <cstdlib>
#include <string>
#include <vector>
#include <unordered_map>
#include "tinyxml2.h"
#include "cpp_pptx_template.h"

namespace cpp_pptx_template {

	struct slide_template {
		tinyxml2::XMLDocument doc_slide, doc_slide_rel;
		struct element_info {
			tinyxml2::XMLElement *elm, *elm_off, *elm_ext;
			int off_x, off_y, ext_cx, ext_cy;
			static const char *find_name(tinyxml2::XMLElement *elm_, const char *nvName) {
				tinyxml2::XMLElement *celm[2];
				if (!(celm[0] = elm_->FirstChildElement(nvName))) return 0;
				if (!(celm[1] = celm[0]->FirstChildElement("p:cNvPr"))) return 0;
				const char *name = 0;
				if ((celm[1]->QueryAttribute("name", &name), name) == 0) return 0;
				return name;
			}
			bool read(tinyxml2::XMLElement *elm_) {
				elm = elm_;
				tinyxml2::XMLElement *celm[2];
				if (!(celm[0] = elm->FirstChildElement("p:spPr"))) return false;
				if (!(celm[1] = celm[0]->FirstChildElement("a:xfrm"))) return false;

				const char *x_str = 0, *y_str = 0, *cx_str = 0, *cy_str = 0;
				if (!(elm_off = celm[1]->FirstChildElement("a:off"))) return false;
				if ((elm_off->QueryAttribute("x", &x_str), x_str) == 0) return false;
				off_x = std::strtol(x_str, 0, 10);
				if ((elm_off->QueryAttribute("y", &y_str), y_str) == 0) return false;
				off_y = std::strtol(y_str, 0, 10);

				if (!(elm_ext = celm[1]->FirstChildElement("a:ext"))) return false;
				if ((elm_ext->QueryAttribute("cx", &cx_str), cx_str) == 0) return false;
				ext_cx = std::strtol(cx_str, 0, 10);
				if ((elm_ext->QueryAttribute("cy", &cy_str), cy_str) == 0) return false;
				ext_cy = std::strtol(cy_str, 0, 10);
				return true;
			}
		};
		struct image_info {
			element_info p_pic;
			double aspect_y_x;
			std::string rId;
		};
		std::unordered_map<std::string, image_info> images;
		std::unordered_map<std::string, tinyxml2::XMLElement *> image_rels; // key:rId, value:relationエレメント
		struct shape_info {
			element_info p_sp;
			tinyxml2::XMLElement *a_r;
		};
		std::unordered_map<std::string, shape_info> shapes;
		bool parse_slide(const char *xml_slide) {
			doc_slide.Clear();
			doc_slide.Parse(xml_slide);
			tinyxml2::XMLElement *elm[8];
			if (!(elm[0] = doc_slide.FirstChildElement("p:sld"))) return false;
			if (!(elm[1] = elm[0]->FirstChildElement("p:cSld"))) return false;
			if (!(elm[2] = elm[1]->FirstChildElement("p:spTree"))) return false;
			elm[3] = elm[2]->FirstChildElement();
			while (elm[3]) {
				const char *elm_name = elm[3]->Name();
				if (std::strcmp(elm_name, "p:pic") == 0) {
					const char *name = element_info::find_name(elm[3], "p:nvPicPr");
					if (!name) return false;
					image_info &ii = images[name];
					if (!ii.p_pic.read(elm[3]) || ii.p_pic.ext_cx == 0) return false;
					ii.aspect_y_x = static_cast<double>(ii.p_pic.ext_cy) / static_cast<double>(ii.p_pic.ext_cx);
					if (!(elm[4] = elm[3]->FirstChildElement("p:blipFill"))) return false;
					if (!(elm[5] = elm[4]->FirstChildElement("a:blip"))) return false;
					const char *r_embed = 0;
					if ((elm[5]->QueryAttribute("r:embed", &r_embed), r_embed) == 0) return false;
					ii.rId = r_embed;
				} else if (std::strcmp(elm_name, "p:sp") == 0) {
					const char *name = element_info::find_name(elm[3], "p:nvSpPr");
					if (!name) return false;
					shape_info &si = shapes[name];
					if (!si.p_sp.read(elm[3])) return false;
					if (!(elm[4] = elm[3]->FirstChildElement("p:txBody"))) return false;
					if (!(elm[5] = elm[4]->FirstChildElement("a:p"))) return false;
					if (!(elm[6] = elm[5]->FirstChildElement("a:r"))) return false;
					si.a_r = elm[6];

					tinyxml2::XMLElement *a_t_dest, *a_t_src;
					if (!(a_t_dest = elm[6]->FirstChildElement("a:t"))) return false;

					while (elm[7] = elm[6]->NextSiblingElement("a:r")) {
						if (!(a_t_src = elm[7]->FirstChildElement("a:t"))) return false;
						a_t_dest->InsertNewText(a_t_src->GetText());
						elm[5]->DeleteChild(elm[7]);
					}
				}
				elm[3] = elm[3]->NextSiblingElement();
			}
			return true;
		}
		bool parse_slide_rel(const char *xml_slide_rel) {
			doc_slide_rel.Clear();
			doc_slide_rel.Parse(xml_slide_rel);
			tinyxml2::XMLElement *elm[2];
			if (!(elm[0] = doc_slide_rel.FirstChildElement("Relationships"))) return false;
			if (!(elm[1] = elm[0]->FirstChildElement("Relationship"))) return false;
			while (elm[1]) {
				const char *idname = 0;
				if ((elm[1]->QueryAttribute("Id", &idname), idname) == 0) return false;
				image_rels[idname] = elm[1];
				elm[1] = elm[1]->NextSiblingElement("Relationship");
			}
			return true;
		}
		const char *slide_to_string(tinyxml2::XMLPrinter &printer) {
			printer.ClearBuffer(), doc_slide.Print(&printer);
			return printer.CStr();
		}
		const char *slide_rel_to_string(tinyxml2::XMLPrinter &printer) {
			printer.ClearBuffer(), doc_slide_rel.Print(&printer);
			return printer.CStr();
		}
		bool change_shape_text(const char *name, const char *text) {
			auto iter = shapes.find(name);
			if (iter == shapes.end()) return false;
			auto &si = iter->second;

			tinyxml2::XMLElement *a_t;
			if (!(a_t = si.a_r->FirstChildElement("a:t"))) return false;
			a_t->SetText(text);
			return true;
		}
		bool set_image_data(const char *name, int image_index, const image_data &image_item, bool omit_export_img, zipper::Zip &zip) {
			auto iter = images.find(name);
			if (iter == images.end()) return false;
			auto &ii = iter->second;
			if (!ii.p_pic.elm_off || !ii.p_pic.elm_ext) return false;
			auto iter2 = image_rels.find(ii.rId);
			if (iter2 == image_rels.end()) return false;
			auto rel_elm = iter2->second;

			char buf[100];

			int img_x, img_y, img_cx, img_cy;
			if (ii.aspect_y_x > image_item.aspect_y_x) { // 枠よりも画像の方が横長の時
				img_cx = ii.p_pic.ext_cx;
				img_x = ii.p_pic.off_x;
				img_cy = static_cast<int>(std::floor(static_cast<double>(ii.p_pic.ext_cx) * image_item.aspect_y_x + 0.5));
				img_y = static_cast<int>(std::floor(static_cast<double>(ii.p_pic.off_y) + (ii.p_pic.ext_cy - img_cy) / 2 + 0.5));
			} else { // 枠よりも画像の方が縦長の時
				img_cy = ii.p_pic.ext_cy;
				img_y = ii.p_pic.off_y;
				img_cx = static_cast<int>(std::floor(static_cast<double>(ii.p_pic.ext_cy) / image_item.aspect_y_x + 0.5));
				img_x = static_cast<int>(std::floor(static_cast<double>(ii.p_pic.off_x) + (ii.p_pic.ext_cx - img_cx) / 2 + 0.5));
			}
			std::snprintf(buf, 99, "%d", img_x), ii.p_pic.elm_off->SetAttribute("x", buf);
			std::snprintf(buf, 99, "%d", img_y), ii.p_pic.elm_off->SetAttribute("y", buf);
			std::snprintf(buf, 99, "%d", img_cx), ii.p_pic.elm_ext->SetAttribute("cx", buf);
			std::snprintf(buf, 99, "%d", img_cy), ii.p_pic.elm_ext->SetAttribute("cy", buf);

			std::snprintf(buf, 99, "media/image%d.%s", image_index + 1, image_item.ext);
			rel_elm->SetAttribute("Target", (std::string("../") + buf).c_str());
			if (!omit_export_img) zip.add_file((std::string("ppt/") + buf), image_item.filedata, image_item.filedata_size);
			return true;
		}
	};

	void create_slide_from_template(zipper::UnZip &template_pptx, zipper::Zip &pptx_to_generate, const nlohmann::json &slides_define, const image_data *image_data_ary, int image_data_count) {
		std::unordered_map<int, slide_template> slide_templates;
		std::string cont_type, presentation_xml, presentation_xml_rels;

		// スライド関連のxmlファイルを読み込む
		// 関連ファイル以外はそのまま出力ファイルに書き出す。
		template_pptx.enumerate([&](auto &unzip) {
			auto file_path = template_pptx.file_path();
			const char *s = file_path.c_str();

			// スライドのxmlファイル内容を退避
			if (std::strstr(s, "ppt/slides/") == s) {
				std::string xml;
				template_pptx.read(xml);
				if (std::strstr(s + 11, "_rels/slide") == s + 11) {
					int idx = std::strtol(s + 11 + 11, 0, 10);
					slide_templates[idx].parse_slide_rel(xml.c_str());
				} else if (std::strstr(s + 11, "slide") == s + 11) {
					int idx = std::strtol(s + 11 + 5, 0, 10);
					slide_templates[idx].parse_slide(xml.c_str());
				}
			} else if (std::strstr(s, "[Content_Types].xml") == s) {
				template_pptx.read(cont_type);
			} else if (std::strstr(s, "ppt/presentation.xml") == s) {
				template_pptx.read(presentation_xml);
			} else if (std::strstr(s, "ppt/_rels/presentation.xml.rels") == s) {
				template_pptx.read(presentation_xml_rels);
			} else if (template_pptx.is_dir()) {
				pptx_to_generate.add_dir(s);
			} else if (std::strstr(s, "ppt/media/") != s) {
				std::string cont;
				template_pptx.read(cont);
				pptx_to_generate.add_file(s, cont);
			}
		});

		tinyxml2::XMLDocument doc;
		tinyxml2::XMLPrinter printer(0, false);
		char buf[100];

		doc.Clear(), doc.Parse(cont_type.c_str());
		do {
			tinyxml2::XMLElement *elm[3], *elmSlideToInsert = 0;
			if (!(elm[0] = doc.FirstChildElement("Types"))) break;
			if (!(elm[1] = elm[0]->FirstChildElement("Override"))) break;

			const char *slide_master_cont_type = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
			const char *slide_cont_type = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

			// slide*.xml の型を定義しているエレメントを消す、スライドマスターのエレメントを記憶する
			do {
				elm[2] = elm[1]->NextSiblingElement("Override");
				const char *contType = 0;
				elm[1]->QueryAttribute("ContentType", &contType);
				if (contType) {
					if (std::strcmp(contType, slide_cont_type) == 0) {
						elm[0]->DeleteChild(elm[1]);
					} else if (std::strcmp(contType, slide_master_cont_type) == 0) {
						elmSlideToInsert = elm[1];
					}
				}
				elm[1] = elm[2];
			} while (elm[1]);

			// スライドマスターのエレメントの後にスライドエレメントを足す
			for (size_t i = 0; i < slides_define.size(); ++i) {
				elm[1] = doc.NewElement("Override");
				elm[1]->SetAttribute("ContentType", slide_cont_type);
				if (elmSlideToInsert) elm[0]->InsertAfterChild(elmSlideToInsert, elm[1]);
				else elm[0]->InsertEndChild(elm[1]);
				elmSlideToInsert = elm[1];
				std::snprintf(buf, 99, "/ppt/slides/slide%zu", i + 1); elm[1]->SetAttribute("PartName", buf);
			}

			printer.ClearBuffer(), doc.Print(&printer);
			pptx_to_generate.add_file("[Content_Types].xml", printer.CStr());
		} while (0);

		doc.Clear(), doc.Parse(presentation_xml.c_str());
		do {
			// スライドリストの内容を一度消して、スライド枚数分、エレメントを挿入する
			tinyxml2::XMLElement *elm[3];
			if (!(elm[0] = doc.FirstChildElement("p:presentation"))) break;
			if (!(elm[1] = elm[0]->FirstChildElement("p:sldIdLst"))) break;
			elm[1]->DeleteChildren();
			for (size_t i = 0; i < slides_define.size(); ++i) {
				elm[2] = elm[1]->InsertNewChildElement("p:sldId");
				std::snprintf(buf, 99, "rId%zu", i + 2); elm[2]->SetAttribute("r:id", buf);
				std::snprintf(buf, 99, "%zu", i + 256); elm[2]->SetAttribute("id", buf);
			}

			printer.ClearBuffer(), doc.Print(&printer);
			pptx_to_generate.add_file("ppt/presentation.xml", printer.CStr());
		} while (0);

		doc.Clear(), doc.Parse(presentation_xml_rels.c_str());
		do {
			// テンプレートのスライドのリレーションシップエレメントの数を取得する
			tinyxml2::XMLElement *elm[3], *relFirst;
			if (!(elm[0] = doc.FirstChildElement("Relationships"))) break;
			relFirst = elm[1] = elm[0]->FirstChildElement("Relationship");
			int template_slide_cnt = 0;
			const char *slidemaster_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
			const char *slide_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
			while (elm[1]) {
				const char *rel_type = 0;
				elm[1]->QueryAttribute("Type", &rel_type);
				elm[2] = elm[1]->NextSiblingElement("Relationship");
				if (rel_type && std::strcmp(rel_type, slide_rel_type) == 0) {
					elm[0]->DeleteChild(elm[1]);
					++template_slide_cnt;
				}
				elm[1] = elm[2];
			}

			// スライドマスタ以外のリレーションシップエレメントのIdを、(作成するスライド数 - テンプレートスライド数) だけ増分する
			// スライドマスタのIdの最大値を取得する
			int slide_incr = static_cast<int>(slides_define.size()) - template_slide_cnt;
			int slide_master_id_max = 0;
			elm[1] = relFirst;
			while (elm[1]) {
				const char *rel_type = 0, *id = 0;
				elm[1]->QueryAttribute("Type", &rel_type);
				if (rel_type) {
					elm[1]->QueryAttribute("Id", &id);
					if (!id && std::strstr(id, "rId") != id) continue;
					int id_cur = std::strtol(id + 3, 0, 10);
					if (std::strcmp(rel_type, slidemaster_rel_type) != 0) {
						std::snprintf(buf, 99, "rId%d", id_cur + slide_incr); elm[1]->SetAttribute("Id", buf);
					} else {
						if (slide_master_id_max < id_cur) slide_master_id_max = id_cur;
					}
				}
				elm[1] = elm[1]->NextSiblingElement("Relationship");
			}

			// スライドマスタのリレーションシップエレメントを足す
			for (int i = 0; i < static_cast<int>(slides_define.size()); ++i) {
				elm[1] = elm[0]->InsertNewChildElement("Relationship");
				std::snprintf(buf, 99, "rId%d", slide_master_id_max + i + 1); elm[1]->SetAttribute("Id", buf);
				elm[1]->SetAttribute("Type", slide_rel_type);
				std::snprintf(buf, 99, "slides/slide%d.xml", i + 1); elm[1]->SetAttribute("Target", buf);
			}

			printer.ClearBuffer(), doc.Print(&printer);
			pptx_to_generate.add_file("ppt/_rels/presentation.xml.rels", printer.CStr());
		} while (0);

		std::vector<int> image_exported;
		image_exported.resize(image_data_count);
		for (size_t i = 0; i < slides_define.size(); ++i) {
			auto &slide_define = slides_define[i];
			int slide_template_index = slide_define["template_slide"] + 1;
			auto &cur_template = slide_templates[slide_template_index];

			auto &slide_shapes = slide_define["shapes"];
			for (auto iter = slide_shapes.begin(); iter != slide_shapes.end(); ++iter) {
				auto &val = iter.value();
				const auto &text_p = val.find("text");
				if (text_p == val.end() && text_p.value().is_string()) continue;
				const std::string &shape_text = text_p.value();
				cur_template.change_shape_text(iter.key().c_str(), shape_text.c_str());
			}

			auto &slide_images = slide_define["images"];
			for (auto iter = slide_images.begin(); iter != slide_images.end(); ++iter) {
				auto &val = iter.value();
				const auto &image_data_p = val.find("image_data");
				if (image_data_p == val.end() && !image_data_p.value().is_number_integer()) continue;
				int image_data_index = image_data_p.value();
				cur_template.set_image_data(iter.key().c_str(), image_data_index, image_data_ary[image_data_index], image_exported[image_data_index], pptx_to_generate);
				image_exported[image_data_index] = 1;
			}
			std::snprintf(buf, 99, "ppt/slides/slide%zu.xml", i + 1);
			pptx_to_generate.add_file(buf, cur_template.slide_to_string(printer));
			std::snprintf(buf, 99, "ppt/slides/_rels/slide%zu.xml.rels", i + 1);
			pptx_to_generate.add_file(buf, cur_template.slide_rel_to_string(printer));
		}
	}
}
