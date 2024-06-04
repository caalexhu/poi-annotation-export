package org.example.model.vo;

import lombok.Data;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.example.util.PoiExportUtil;

/**
 * @author: Alex Hu
 * @createTime: 2024/06/03 18:57
 * @description:
 */
@Data
public class OrderVO {
    @PoiExportUtil.PoiExportField(label = "订单编号", order = 1, align = HorizontalAlignment.CENTER)
    private String orderNo;
    @PoiExportUtil.PoiExportField(label = "订单用户", order = 2, align = HorizontalAlignment.CENTER)
    private String orderUser;
    @PoiExportUtil.PoiExportField(label = "订单时间", order = 3, align = HorizontalAlignment.CENTER)
    private String orderTime;
    @PoiExportUtil.PoiExportField(label = "订单金额", order = 4, width = 15, align = HorizontalAlignment.RIGHT)
    private String orderAmount;
    private String orderDesc;
    private String orderRemark;
    private String orderPhone;
    private String orderZipCode;
    @PoiExportUtil.PoiExportField(label = "订单国家", subGroup = true)
    private String orderCountry;
    @PoiExportUtil.PoiExportField(label = "订单省份", subGroup = true)
    private String orderProvince;
    @PoiExportUtil.PoiExportField(label = "订单城市", order = 6)
    private String orderCity;
    @PoiExportUtil.PoiExportField(label = "详细地址", order = 7)
    private String orderAddressDetail;

}
