#pragma once
#include <string>

class Trie {
private:
    bool isLeaf;        
    Trie* children[26];
public:
    Trie() {
        this->isLeaf = false;
        for (int i = 0; i < 26; i++) {
            this->children[i] = nullptr;
        }
    }
    
    void insert(std::string);
    bool search(std::string);
};
