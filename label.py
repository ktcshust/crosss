import torch
import torch.nn as nn
import torch.optim as optim
import torch.nn.functional as F
import pandas as pd
import numpy as np
import random
from sklearn.metrics import accuracy_score, precision_score, recall_score, f1_score
from copy import deepcopy

# -------------------------------
# CONFIGURABLE HYPERPARAMETERS
# -------------------------------
embedding_dim = 128  # Kích thước vector embedding (sẽ được tối ưu bằng EA)
hidden_dim = 64  # Kích thước tầng ẩn của Siamese network
initial_margin = 1.0  # Margin ban đầu cho hàm contrastive loss
learning_rate = 1e-3  # Learning rate
num_epochs = 10  # Số epoch huấn luyện


# -------------------------------
# DATA LOADING TỪ EXCEL
# -------------------------------
def get_text_from_excel(file_path):
    """
    Đọc file Excel và chuyển toàn bộ nội dung thành 1 chuỗi.
    Bạn có thể tùy chỉnh cách xử lý này dựa trên cấu trúc file.
    """
    try:
        df = pd.read_excel(file_path)
        # Thay thế giá trị NaN bằng chuỗi rỗng
        df = df.fillna('')
        # Nối tất cả các cell thành một chuỗi
        text = " ".join(df.astype(str).values.flatten().tolist())
        return text
    except Exception as e:
        print(f"Lỗi đọc file {file_path}: {e}")
        return ""


def load_dataset(metadata_csv):
    """
    Đọc file CSV metadata có các cột: id, label, file_full, file_key
    Trả về danh sách các record dạng dict.
    """
    df = pd.read_csv(metadata_csv)
    data = []
    for idx, row in df.iterrows():
        record = {
            "id": row["id"],
            "label": int(row["label"]),
            "file_full_text": get_text_from_excel(row["file_full"]),
            "file_key_text": get_text_from_excel(row["file_key"])
        }
        data.append(record)
    return data


# Giả sử bạn có 2 file metadata:
train_metadata_path = "metadata.csv"  # Tập train: mỗi label có 1 record
#test_metadata_path = "metadata_test.csv"  # Tập test: mỗi label có 4 record

train_data = load_dataset(train_metadata_path)
#test_data = load_dataset(test_metadata_path)


# -------------------------------
# PLACEHOLDER: GET EMBEDDING
# -------------------------------
def get_embedding(text):
    """
    Thay thế hàm này bằng lời gọi API OpenAI Embedding nếu cần.
    Tạm thời dùng vector ngẫu nhiên dựa trên seed của text để đảm bảo nhất quán.
    """
    seed = hash(text) % (2 ** 32)
    rng = np.random.default_rng(seed)
    emb = rng.standard_normal(embedding_dim)
    return torch.tensor(emb, dtype=torch.float)


# -------------------------------
# SIAMESE NETWORK ARCHITECTURE
# -------------------------------
class SiameseNetwork(nn.Module):
    def __init__(self, input_dim, hidden_dim, output_dim):
        super(SiameseNetwork, self).__init__()
        self.fc1 = nn.Linear(input_dim, hidden_dim)
        self.fc2 = nn.Linear(hidden_dim, output_dim)

    def forward_once(self, x):
        out = F.relu(self.fc1(x))
        out = self.fc2(out)
        return out

    def forward(self, emb_full, emb_key):
        out_full = self.forward_once(emb_full)
        out_key = self.forward_once(emb_key)
        return out_full, out_key


# -------------------------------
# CONTRASTIVE LOSS FUNCTION
# -------------------------------
def contrastive_loss(out1, out2, label, margin):
    """
    label = 1: positive pair (các embedding nên gần nhau)
    label = 0: negative pair (các embedding nên cách xa nhau hơn margin)
    """
    euclidean_distance = F.pairwise_distance(out1, out2)
    loss = torch.mean(label * torch.pow(euclidean_distance, 2) +
                      (1 - label) * torch.pow(torch.clamp(margin - euclidean_distance, min=0.0), 2))
    return loss


# -------------------------------
# RL AGENT ĐIỀU CHỈNH MARGIN
# -------------------------------
class RLAgent:
    def __init__(self, margin):
        self.margin = margin
        self.step = 0.05

    def update(self, prev_loss, current_loss):
        if current_loss < prev_loss:
            self.margin += self.step
        else:
            self.margin = max(0.1, self.margin - self.step)
        return self.margin


# -------------------------------
# TRAINING LOOP VỚI RL
# -------------------------------
def train_model(model, train_data, num_epochs, learning_rate, initial_margin):
    optimizer = optim.Adam(model.parameters(), lr=learning_rate)
    rl_agent = RLAgent(initial_margin)
    current_margin = initial_margin
    prev_epoch_loss = None

    # Với one-shot: mỗi record của train_data là duy nhất cho 1 label
    for epoch in range(num_epochs):
        total_loss = 0.0
        for record in train_data:
            # Positive pair: file_full và file_key của record đó
            emb_full = get_embedding(record["file_full_text"]).unsqueeze(0)
            emb_key = get_embedding(record["file_key_text"]).unsqueeze(0)
            out_full, out_key = model(emb_full, emb_key)
            label_pos = torch.tensor([1.0])
            loss_pos = contrastive_loss(out_full, out_key, label_pos, current_margin)

            # Negative pair: ghép file_full của record với file_key của record có label khác
            neg_candidates = [r for r in train_data if r["label"] != record["label"]]
            if neg_candidates:
                neg_record = random.choice(neg_candidates)
                emb_key_neg = get_embedding(neg_record["file_key_text"]).unsqueeze(0)
                out_full_neg, out_key_neg = model(emb_full, emb_key_neg)
                label_neg = torch.tensor([0.0])
                loss_neg = contrastive_loss(out_full_neg, out_key_neg, label_neg, current_margin)
            else:
                loss_neg = 0.0

            loss = loss_pos + loss_neg

            optimizer.zero_grad()
            loss.backward()
            optimizer.step()

            total_loss += loss.item()
        avg_loss = total_loss / len(train_data)
        if prev_epoch_loss is not None:
            current_margin = rl_agent.update(prev_epoch_loss, avg_loss)
        prev_epoch_loss = avg_loss
        print(f"Epoch {epoch + 1}/{num_epochs} - Avg Loss: {avg_loss:.4f} - Margin: {current_margin:.4f}")
    return model, current_margin


# -------------------------------
# EA: EVOLUTIONARY ALGORITHM CHO HYPERPARAMETER OPTIMIZATION
# -------------------------------
def evaluate_fitness(hyperparams, train_data):
    global embedding_dim
    embedding_dim = hyperparams["embedding_dim"]

    model = SiameseNetwork(input_dim=embedding_dim, hidden_dim=hidden_dim, output_dim=embedding_dim)
    # Huấn luyện nhanh qua vài epoch để đánh giá fitness
    model, final_margin = train_model(model, train_data, num_epochs=3,
                                      learning_rate=hyperparams["lr"],
                                      initial_margin=hyperparams["margin"])
    # Ví dụ: fitness được định nghĩa dựa trên margin cuối (giả sử margin càng cao mô hình phân biệt tốt hơn)
    return final_margin


def ea_optimization(train_data, generations=3, population_size=4):
    population = []
    best_individual = None
    for _ in range(population_size):
        individual = {
            "lr": random.choice([1e-3, 5e-4, 1e-4]),
            "margin": random.uniform(0.5, 1.5),
            "embedding_dim": random.choice([64, 128, 256])
        }
        population.append(individual)

    for gen in range(generations):
        print(f"--- Generation {gen + 1} ---")
        fitness_scores = []
        for individual in population:
            fitness = evaluate_fitness(individual, train_data)
            fitness_scores.append(fitness)
            print(f"Individual {individual} => Fitness: {fitness:.4f}")
        best_idx = np.argmax(fitness_scores)
        best_individual = population[best_idx]
        print(f"Best Individual: {best_individual} with fitness {fitness_scores[best_idx]:.4f}")

        new_population = [best_individual]
        while len(new_population) < population_size:
            child = deepcopy(best_individual)
            if random.random() < 0.5:
                child["lr"] *= random.uniform(0.8, 1.2)
            if random.random() < 0.5:
                child["margin"] *= random.uniform(0.8, 1.2)
            if random.random() < 0.5:
                child["embedding_dim"] = random.choice([64, 128, 256])
            new_population.append(child)
        population = new_population
    return best_individual


# -------------------------------
# EVALUATION METRICS TRÊN TẬP TEST
# -------------------------------
#def evaluate_model(model, test_data, margin):
    #model.eval()
    #y_true = []
    #y_pred = []
    #cosine_similarities = []

    #with torch.no_grad():
        #for record in test_data:
            #emb_full = get_embedding(record["file_full_text"]).unsqueeze(0)
            #emb_key = get_embedding(record["file_key_text"]).unsqueeze(0)
            #out_full, out_key = model(emb_full, emb_key)

            #cos_sim = F.cosine_similarity(out_full, out_key)
            #cosine_similarities.append(cos_sim.item())
            # Ở đây, giả sử các cặp test là positive pairs nên label thật luôn là 1.
            #pred_label = 1 if cos_sim.item() > 0.5 else 0  # Ngưỡng điều chỉnh theo bài toán
            #y_pred.append(pred_label)
            #y_true.append(1)

    #accuracy = accuracy_score(y_true, y_pred)
    #precision = precision_score(y_true, y_pred, zero_division=0)
    #recall = recall_score(y_true, y_pred, zero_division=0)
    #f1 = f1_score(y_true, y_pred, zero_division=0)
    #avg_cos_sim = np.mean(cosine_similarities)

    #print(f"Test Accuracy: {accuracy:.4f}")
    #print(f"Precision: {precision:.4f} | Recall: {recall:.4f} | F1-Score: {f1:.4f}")
    #print(f"Average Cosine Similarity: {avg_cos_sim:.4f}")

    #return accuracy, precision, recall, f1, avg_cos_sim


# -------------------------------
# MAIN FUNCTION: CHU TRÌNH HOÀN CHỈNH
# -------------------------------
def main():
    # EA Optimization: Tìm hyperparameter tối ưu
    best_hyperparams = ea_optimization(train_data, generations=3, population_size=4)
    print("Optimized Hyperparameters:", best_hyperparams)

    global embedding_dim
    embedding_dim = best_hyperparams["embedding_dim"]

    model = SiameseNetwork(input_dim=embedding_dim, hidden_dim=hidden_dim, output_dim=embedding_dim)

    # Huấn luyện model với tích hợp RL trên tập train (one-shot: 1 record cho mỗi label)
    model, final_margin = train_model(model, train_data, num_epochs=num_epochs,
                                      learning_rate=best_hyperparams["lr"],
                                      initial_margin=best_hyperparams["margin"])

    # Đánh giá trên tập test (mỗi label có 4 record)
    #evaluate_model(model, test_data, final_margin)


if __name__ == "__main__":
    main()
