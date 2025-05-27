from sentence_transformers import SentenceTransformer, util
import nltk
import re

nltk.download('punkt')

def segment_english_text(text):
    sentences = nltk.sent_tokenize(text)
    return sentences

def segment_chinese_text(text):
    pattern = re.compile(r'[^。！？]*[。！？]?')
    sentences = [s.strip() for s in pattern.findall(text) if s.strip()]
    return sentences

model = SentenceTransformer('sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2')

def encode_sentences(sentences):
    return model.encode(sentences, convert_to_tensor=True)

def align_changes(english_changes, english_full_text, chinese_full_text):
    eng_sentences = segment_english_text(english_full_text)
    ch_sentences = segment_chinese_text(chinese_full_text)

    eng_embeds = encode_sentences(eng_sentences)
    ch_embeds = encode_sentences(ch_sentences)

    mapped_changes = []

    for change in english_changes:
        c_text = change['text']
        c_type = change['type']

        # Embed changed text itself (smaller chunk)
        change_embed = model.encode(c_text, convert_to_tensor=True)

        # Find closest English sentence to change_text embedding (in case it isn't a full sentence)
        cos_scores_eng = util.pytorch_cos_sim(change_embed, eng_embeds)[0]
        best_eng_idx = cos_scores_eng.argmax().item()

        # Now find best matching Chinese sentence to that English sentence embedding
        query_eng_embed = eng_embeds[best_eng_idx]
        cos_scores_ch = util.pytorch_cos_sim(query_eng_embed, ch_embeds)[0]
        best_ch_idx = cos_scores_ch.argmax().item()

        mapped_changes.append({
            'type': c_type,
            'english_text': c_text,
            'chinese_text': ch_sentences[best_ch_idx]
        })

    return mapped_changes
