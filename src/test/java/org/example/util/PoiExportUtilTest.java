package org.example.util;


import org.example.model.vo.OrderVO;
import org.junit.Test;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Random;
import java.util.stream.Collectors;

/**
 * @author: Alex Hu
 * @createTime: 2024/06/03 19:16
 * @description:
 */

public class PoiExportUtilTest {


    @Test
    public void exportToExcel() {
        PoiExportUtil poiExportUtil = new PoiExportUtil();
        List<OrderVO> orderVOList = generateOrders();
        poiExportUtil.exportToExcel(orderVOList, "order.xlsx");
    }

    public List<OrderVO> generateOrders() {
        String[] COUNTRIES = {"China", "Japan", "Canada"};
        Random RANDOM = new Random();
        List<OrderVO> orders = new ArrayList<>();
        for (int i = 0; i < 30; i++) {
            String orderNo = "OrderNo" + (i + 1);
            String orderUser = "User" + (i + 1);
            String orderTime = "Time" + (i + 1);
            String orderAmount = RANDOM.nextInt(10000) + ".00";
            String orderDesc = "Desc" + (i + 1);
            String orderRemark = "Remark" + (i + 1);
            String orderPhone = "Phone" + (i + 1);
            String orderZipCode = "ZipCode" + (i + 1);
            String orderCountry = COUNTRIES[RANDOM.nextInt(COUNTRIES.length)];
            String orderProvince = "Province" + (i + 1) % 3;
            String orderCity = "City" + (i + 1);
            String orderAddressDetail = "AddressDetail" + (i + 1);
            OrderVO order = createOrder(orderNo, orderUser, orderTime, orderAmount, orderDesc, orderRemark, orderPhone,
                    orderZipCode, orderCountry, orderProvince, orderCity, orderAddressDetail);
            orders.add(order);
        }
        // Sort by orderCountry and orderTime
        return orders.stream()
                .sorted(Comparator.comparing(OrderVO::getOrderCountry)
                        .thenComparing(OrderVO::getOrderProvince)
                        .thenComparing(OrderVO::getOrderTime))
                .collect(Collectors.toList());
    }


    private OrderVO createOrder(String orderNo, String orderUser, String orderTime, String orderAmount,
                                String orderDesc, String orderRemark, String orderPhone, String orderZipCode,
                                String orderCountry, String orderProvince, String orderCity, String orderAddressDetail) {
        OrderVO order = new OrderVO();
        order.setOrderNo(orderNo);
        order.setOrderUser(orderUser);
        order.setOrderTime(orderTime);
        order.setOrderAmount(orderAmount);
        order.setOrderDesc(orderDesc);
        order.setOrderRemark(orderRemark);
        order.setOrderPhone(orderPhone);
        order.setOrderZipCode(orderZipCode);
        order.setOrderCountry(orderCountry);
        order.setOrderProvince(orderProvince);
        order.setOrderCity(orderCity);
        order.setOrderAddressDetail(orderAddressDetail);
        return order;
    }
}
