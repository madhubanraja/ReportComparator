package com.aisquad.valueadd.reportcompare.pdf;

public interface ResultCollector {

    void addPage(PageDiffCalculator diffCalculator, int pageIndex,
            ImageWithDimension expectedImage, ImageWithDimension actualImage, ImageWithDimension diffImage);

    void noPagesFound();

    default void done() {}
}
