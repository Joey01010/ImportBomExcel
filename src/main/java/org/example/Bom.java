package org.example;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class Bom<T> implements Comparable {
    private String BomId;
    private List<T> items;
    private String circuit;

    public Bom() {
    }

    public Bom(String bomId, List<T> items, String circuit) {
        BomId = bomId;
        this.items = items;
        this.circuit = circuit;
    }

    public String getBomId() {
        return BomId;
    }

    public void setBomId(String bomId) {
        BomId = bomId;
    }

    public List<T> getItems() {
        if (this.items == null) {
            this.items = new ArrayList();
        }
        return items;
    }

    public void setItems(List<T> items) {
        this.items = items;
    }

    public String getCircuit() {
        return circuit;
    }

    public void setCircuit(String circuit) {
        this.circuit = circuit;
    }

    @Override
    public String toString() {
        return "Bom{" +
                "BomId='" + BomId + '\'' +
                ", items=" + items +
                ", circuit='" + circuit + '\'' +
                '}';
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Bom<?> bom = (Bom<?>) o;
        return BomId.equals(bom.BomId);
    }

    @Override
    public int hashCode() {
        return Objects.hash(BomId);
    }

    @Override
    public int compareTo(Object o) {
        if (o instanceof Bom) {
            Bom bom = (Bom) o;
            return this.BomId.compareTo(bom.getBomId());
        }
        throw new RuntimeException("传入的数据类型不一致！");
    }
}
