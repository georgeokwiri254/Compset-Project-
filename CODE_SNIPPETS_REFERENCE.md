# Exact Code Snippets from Hotel_Similarity_Analysis_Enhanced.ipynb

## 1. Feature Weights Definition
**Location: Cell 11**

```python
feature_weights = {
    # TIER 1: Critical Success Factors (2.0 - 2.4)
    'Star\nPoints': 2.4,                    # Hotel classification - HIGHEST priority
    'TripAdvisor\nPoints': 2.0,            # Guest satisfaction & reputation
    'Booking.com\nPoints': 2.0,            # Booking conversion & ratings
    'Total Keys\nPoints': 2.0,             # Hotel capacity/scale
    'Meeting\nPoints': 2.0,                # MICE segment capability
    'F&B\nPoints': 2.0,                    # Guest experience & revenue
    
    # TIER 2: High Impact Factors (1.5 - 1.7)
    'Opening\nPoints': 1.7,                # Property age/newness
    'Distance\nPoints': 1.5,               # Location proximity
    
    # TIER 3: Moderate Impact (1.2 - 1.3)
    'Room Mix\nPoints (3)': 1.3,           # Room variety
    'Renovation\nPoints': 1.2,             # Recent updates
    
    # TIER 4: Basic Amenities (0.4 - 0.5)
    'Pool\n(1/0)': 0.5,                    # Standard amenity
    'Gym\n(1/0)': 0.5,                     # Standard amenity
    'Spa\n(1/0)': 0.5,                     # Standard amenity
    'Sauna\n(1/0)': 0.5,                   # Standard amenity
    'Kids Club\n(1/0)': 0.4                # Family segment differentiator
}
```

---

## 2. Weights Applied to DataFrame
**Location: Cell 11**

```python
# Create weighted feature matrix
df_weighted = df_features.copy()
for feature, weight in feature_weights.items():
    df_weighted[feature] = df_features[feature] * weight
```

---

## 3. X_weighted Creation (PCA Path)
**Location: Cell 24**

```python
# PCA using WEIGHTED features (not standardized)
X_weighted_array = df_weighted[all_features].values
pca = PCA()
X_pca = pca.fit_transform(X_weighted_array)
```

**Note:** This creates a SEPARATE variable for PCA analysis only.

---

## 4. X_weighted Creation (Main Pipeline)
**Location: Cell 32**

```python
# Calculate cosine similarity using WEIGHTED features
X_weighted = df_weighted[all_features].values
cosine_sim_matrix = cosine_similarity(X_weighted)

# Get similarity scores for Grand Millennium Dubai
similarity_scores = cosine_sim_matrix[target_idx]

# Create results dataframe
similarity_results = pd.DataFrame({
    'Hotel': df['Hotel'],
    'Cosine_Similarity': similarity_scores,
    'Similarity_Percentage': similarity_scores * 100,
    'Overall_Score': df['Normalized\nScore (0-100)']
})
```

**Key Point:** X_weighted is extracted from df_weighted which already contains weighted values.

---

## 5. KNN Model - Weighted Features
**Location: Cell 35**

```python
# CORRECTED: Use weighted features directly without StandardScaler
# This preserves your business-driven feature importance

k_neighbors = 10
knn_model = NearestNeighbors(n_neighbors=k_neighbors+1, metric='cosine')
knn_model.fit(X_weighted)  # ✅ Using weighted features directly

# Find neighbors for Grand Millennium Dubai
target_features = X_weighted[target_idx].reshape(1, -1)
distances, indices = knn_model.kneighbors(target_features)

# Remove the target hotel itself (distance = 0)
distances = distances[0][1:]
indices = indices[0][1:]

# Convert cosine distance to similarity
cosine_similarities_knn = 1 - distances

# Create results dataframe
knn_results = pd.DataFrame({
    'Rank': range(1, k_neighbors + 1),
    'Hotel': df.iloc[indices]['Hotel'].values,
    'Cosine_Distance': distances,
    'Cosine_Similarity': cosine_similarities_knn,
    'Similarity_Percentage': cosine_similarities_knn * 100,
    'Overall_Score': df.iloc[indices]['Normalized\nScore (0-100)'].values
})

print("\n" + "="*100)
print(f"{'K-NEAREST NEIGHBORS ANALYSIS (WITH FEATURE WEIGHTS - CORRECTED)':^100}")
print(f"{'Top 10 Nearest Neighbors to ' + target_hotel:^100}")
print("="*100)
print("\n✅ CORRECTED: Now using weighted features WITHOUT StandardScaler")
print("   Your business-driven weights are FULLY RESPECTED in this analysis\n")
```

**Critical Observations:**
- Uses `X_weighted` directly
- NO StandardScaler applied
- Uses cosine metric (preserves weights)
- Explicit comment confirms this is the corrected approach

---

## 6. Hierarchical Clustering - Ward Linkage
**Location: Cell 39-40**

```python
# CORRECTED: Use weighted features for hierarchical clustering
# This ensures clusters respect your business priorities

linkage_ward = linkage(X_weighted, method='ward')

# Plot dendrogram
plt.figure(figsize=(18, 10))
dendrogram(linkage_ward, labels=df['Hotel'].values, leaf_font_size=11, leaf_rotation=90)
plt.title('Hierarchical Clustering Dendrogram (Ward Linkage) - WITH WEIGHTS', 
         fontsize=16, fontweight='bold', pad=20)
plt.xlabel('Hotel', fontsize=13, fontweight='bold')
plt.ylabel('Euclidean Distance', fontsize=13, fontweight='bold')
plt.tight_layout()
plt.show()

# Perform clustering
n_clusters_ward = 4
ward_clustering = AgglomerativeClustering(n_clusters=n_clusters_ward, linkage='ward')
ward_labels = ward_clustering.fit_predict(X_weighted)  # ✅ Using weighted features

# Add cluster labels to dataframe
df['Ward_Cluster'] = ward_labels

# Find which cluster Grand Millennium Dubai belongs to
target_cluster = df.loc[target_idx, 'Ward_Cluster']

print("\n" + "="*100)
print(f"{'HIERARCHICAL CLUSTERING (Ward Linkage) - WITH FEATURE WEIGHTS':^100}")
print("="*100)
```

**Critical Observations:**
- Uses `X_weighted` directly
- NO scaling applied
- Explicit comment confirms weighted features
- Checkmark confirms this is the correct approach

---

## 7. Hierarchical Clustering - Average Linkage
**Location: Cell 42-43**

```python
# Average linkage with weighted features
linkage_avg = linkage(X_weighted, method='average')

# Plot dendrogram
plt.figure(figsize=(18, 10))
dendrogram(linkage_avg, labels=df['Hotel'].values, leaf_font_size=11, leaf_rotation=90)
plt.title('Hierarchical Clustering Dendrogram (Average Linkage) - WITH WEIGHTS', 
         fontsize=16, fontweight='bold', pad=20)
plt.xlabel('Hotel', fontsize=13, fontweight='bold')
plt.ylabel('Average Distance', fontsize=13, fontweight='bold')
plt.axhline(y=5, color='r', linestyle='--', linewidth=2, label='Distance threshold')
plt.legend()
plt.tight_layout()
plt.show()

# Perform clustering
n_clusters_avg = 4
avg_clustering = AgglomerativeClustering(n_clusters=n_clusters_avg, linkage='average')
avg_labels = avg_clustering.fit_predict(X_weighted)  # ✅ Using weighted features

# Add cluster labels
df['Average_Cluster'] = avg_labels

# Find which cluster Grand Millennium Dubai belongs to
target_cluster_avg = df.loc[target_idx, 'Average_Cluster']

print("\n" + "="*100)
print(f"{'HIERARCHICAL CLUSTERING (Average Linkage) - WITH FEATURE WEIGHTS':^100}")
print("="*100)
```

**Critical Observations:**
- Uses `X_weighted` directly
- NO scaling applied
- Explicit comment confirms weighted features
- Checkmark confirms this is the correct approach

---

## 8. Model Evaluation - Using Weighted Features
**Location: Cell 51**

```python
# Evaluate model quality
silhouette_ward = silhouette_score(X_weighted, ward_labels, metric='euclidean')
silhouette_avg = silhouette_score(X_weighted, avg_labels, metric='euclidean')
ari_score = adjusted_rand_score(ward_labels, avg_labels)

print("\n" + "="*100)
print(f"{'MODEL EVALUATION METRICS':^100}")
print("="*100)

print("\n📊 CLUSTERING QUALITY (Silhouette Scores):")
print("-"*100)
print(f"{'Method':<40} {'Silhouette Score':>20} {'Interpretation':<30}")
print("-"*100)
print(f"{'Ward Linkage':<40} {silhouette_ward:>20.4f} {'Moderate' if silhouette_ward > 0.4 else 'Poor':>30}")
print(f"{'Average Linkage':<40} {silhouette_avg:>20.4f} {'Moderate' if silhouette_avg > 0.4 else 'Poor':>30}")
```

**Critical Observations:**
- Evaluation uses `X_weighted` (weighted version)
- Metrics respect the weighted features
- Both silhouette and ARI scores are computed with weights intact

---

## Import Statement (for Reference)
**Location: Cell 3**

```python
from sklearn.preprocessing import StandardScaler, MinMaxScaler
from sklearn.neighbors import NearestNeighbors
from sklearn.metrics.pairwise import cosine_similarity, cosine_distances
from sklearn.cluster import AgglomerativeClustering
from sklearn.decomposition import PCA
```

**Note:** StandardScaler and MinMaxScaler are imported but NEVER used on X_weighted in the KNN/clustering pipeline.

---

## Summary of Code Flow

1. **weights defined** → feature_weights dict (Cell 11)
2. **weights applied** → df_weighted = df_features * weights (Cell 11)
3. **weights extracted** → X_weighted = df_weighted.values (Cell 32)
4. **weights preserved in KNN** → fit(X_weighted) no scaling (Cell 35)
5. **weights preserved in clustering** → fit_predict(X_weighted) no scaling (Cells 40, 43)
6. **weights in evaluation** → metrics use X_weighted (Cell 51)

**Result: Weights flow through the entire pipeline intact.**

